import streamlit as st
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import tempfile
from google import genai

# Set page configuration
st.set_page_config(page_title="AI PowerPoint Generator", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #333;
        margin-bottom: 1rem;
    }
    .info-text {
        font-size: 1rem;
        color: #555;
    }
    .stButton>button {
        background-color: #1E88E5;
        color: white;
        padding: 0.5rem 1rem;
        font-size: 1rem;
        border-radius: 0.3rem;
    }
</style>
""", unsafe_allow_html=True)

# App title
st.markdown("<h1 class='main-header'>AI PowerPoint Generator</h1>", unsafe_allow_html=True)

# Initialize Gemini API
def initialize_gemini_api(api_key):
    return genai.Client(api_key=api_key)

# Function to generate content using Gemini
def generate_content(client, prompt, topic, num_slides):
    complete_prompt = f"""
    Create content for a {num_slides}-slide PowerPoint presentation about '{topic}'. 
    
    {prompt}
    
    Format the response as a JSON string with this structure:
    {{
        "title": "Main presentation title",
        "slides": [
            {{
                "title": "Slide 1 Title",
                "content": ["Bullet point 1", "Bullet point 2", "Bullet point 3"]
            }},
            {{
                "title": "Slide 2 Title",
                "content": ["Bullet point 1", "Bullet point 2", "Bullet point 3"]
            }}
        ]
    }}
    
    Limit to exactly {num_slides} slides (plus a title slide).
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=complete_prompt,
        )
        
        return response.text
    except Exception as e:
        st.error(f"Error generating content: {str(e)}")
        return None

# Function to parse the AI response into structured content
def parse_ai_response(response_text):
    try:
        import json
        # Find the JSON part in the response (it might be wrapped in markdown code blocks)
        json_content = response_text
        if "```json" in response_text:
            json_content = response_text.split("```json")[1].split("```")[0]
        elif "```" in response_text:
            json_content = response_text.split("```")[1].split("```")[0]
            
        # Parse the JSON
        presentation_data = json.loads(json_content)
        return presentation_data
    except Exception as e:
        st.error(f"Error parsing AI response: {str(e)}")
        return None

# Function to create a PowerPoint presentation
def create_powerpoint(presentation_data):
    ppt = Presentation()
    
    # Set slide dimensions (16:9 aspect ratio)
    ppt.slide_width = Inches(10)
    ppt.slide_height = Inches(5.625)
    
    # Create a title slide
    title_slide_layout = ppt.slide_layouts[0]
    title_slide = ppt.slides.add_slide(title_slide_layout)
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1]
    
    title.text = presentation_data["title"]
    subtitle.text = "Generated with AI PowerPoint Generator"
    
    # Apply formatting to title slide
    title.text_frame.paragraphs[0].font.size = Pt(44)
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 85, 170)
    
    # Create content slides
    for slide_data in presentation_data["slides"]:
        content_slide_layout = ppt.slide_layouts[1]  # Layout with title and content
        slide = ppt.slides.add_slide(content_slide_layout)
        
        # Set slide title
        title = slide.shapes.title
        title.text = slide_data["title"]
        title.text_frame.paragraphs[0].font.size = Pt(36)
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 85, 170)
        
        # Add content as bullet points
        content_placeholder = slide.placeholders[1]
        text_frame = content_placeholder.text_frame
        
        for i, bullet_point in enumerate(slide_data["content"]):
            if i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            # Check for bold text using Markdown syntax
            if "**" in bullet_point:
                parts = bullet_point.split("**")
                for j, part in enumerate(parts):
                    run = p.add_run()
                    run.text = part
                    if j % 2 == 1:  # Odd indices are bold
                        run.font.bold = True
            else:
                p.text = bullet_point
            
            p.font.size = Pt(24)
            p.level = 0  # Top level bullet
    
    # Save presentation to BytesIO object
    output = io.BytesIO()
    ppt.save(output)
    output.seek(0)
    
    return output

# Sidebar for API key
with st.sidebar:
    st.markdown("<h2 class='sub-header'>API Settings</h2>", unsafe_allow_html=True)
    api_key = st.text_input("Enter your Gemini API Key:", type="password")
    st.markdown("<p class='info-text'>Your API key is not stored and is only used for this session.</p>", 
                unsafe_allow_html=True)
    
    # Theme selection
    st.markdown("<h2 class='sub-header'>Presentation Style</h2>", unsafe_allow_html=True)
    color_theme = st.selectbox("Color Theme:", 
                             ["Professional Blue", "Vibrant Orange", "Modern Green", "Elegant Purple"])

# Main content area
st.markdown("<h2 class='sub-header'>Presentation Details</h2>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    presentation_topic = st.text_input("Presentation Topic:", placeholder="Enter the main topic...")
    num_slides = st.slider("Number of Slides:", min_value=3, max_value=15, value=5)

with col2:
    additional_instructions = st.text_area("Additional Instructions:", 
                                         placeholder="E.g., 'Focus on business applications' or 'Include statistics'")

# Generate button
if st.button("Generate Presentation"):
    if not presentation_topic:
        st.error("Please enter a presentation topic.")
    elif not api_key:
        st.error("Please enter your Gemini API key.")
    else:
        with st.spinner("Generating your presentation... This may take a moment."):
            try:
                # Initialize the Gemini client
                client = initialize_gemini_api(api_key)
                
                # Generate content with Gemini
                ai_response = generate_content(client, additional_instructions, presentation_topic, num_slides)
                
                if ai_response:
                    # Parse the AI response
                    presentation_data = parse_ai_response(ai_response)
                    
                    if presentation_data:
                        # Create PowerPoint
                        ppt_file = create_powerpoint(presentation_data)
                        
                        # Display success message and download button
                        st.success("Your presentation has been generated successfully!")
                        
                        # Preview of presentation content
                        with st.expander("Preview Presentation Content", expanded=True):
                            st.subheader(presentation_data["title"])
                            
                            for i, slide in enumerate(presentation_data["slides"]):
                                st.write(f"**Slide {i+1}: {slide['title']}**")
                                for point in slide["content"]:
                                    st.markdown(f"- {point}")
                                st.write("---")
                        
                        # Download button
                        st.download_button(
                            label="Download PowerPoint",
                            data=ppt_file,
                            file_name=f"{presentation_topic.replace(' ', '_')}_presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

# Instructions section
with st.expander("How to Use This App"):
    st.markdown("""
    1. Enter your Gemini API key in the sidebar
    2. Specify the presentation topic and number of slides
    3. Add any additional instructions to customize the content
    4. Click "Generate Presentation" and wait for the AI to create your content
    5. Review the generated content in the preview section
    6. Download your PowerPoint file using the button
    
    The application uses Google's Gemini AI to generate content based on your specifications 
    and creates a professionally formatted PowerPoint presentation that you can download and use.
    """)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center'>Created with Streamlit and Gemini AI</p>", unsafe_allow_html=True)