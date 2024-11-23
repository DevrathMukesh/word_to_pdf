import streamlit as st
import requests

# Streamlit UI
st.title('Word to PDF Converter')
st.write("Upload a Word document to convert it to PDF.")

# File upload
uploaded_file = st.file_uploader("Choose a .docx file", type=["docx"])

if uploaded_file:
    # Display the uploaded file name
    st.write(f"Uploaded file: {uploaded_file.name}")

    # Button to trigger upload to backend
    if st.button('Upload and Convert'):
        # Prepare form data to send to the Flask backend
        file_data = {'file': uploaded_file}

        # Send POST request to Flask backend API to handle file upload
        backend_url = 'http://localhost:5000/upload'  # URL for the Flask API endpoint
        response = requests.post(backend_url, files={'file': (uploaded_file.name, uploaded_file, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')})

        if response.status_code == 200:
            response_json = response.json()
            st.success('File successfully uploaded and converted!')
            # Display metadata
            metadata = response_json['metadata']
            st.write(f"**Author:** {metadata['author']}")
            st.write(f"**Title:** {metadata['title']}")
            st.write(f"**Word Count:** {metadata['word_count']}")

            # Provide the download link for the converted PDF
            download_link = response_json['download_link']

            # Streamlit's download button
            st.download_button(
                label="Download PDF", 
                data=requests.get(download_link).content,  # Request the PDF content directly from Flask
                file_name=f"{uploaded_file.name.split('.')[0]}.pdf",  # Name the downloaded file
                mime="application/pdf"  # MIME type for PDF files
            )

        else:
            st.error('Error during file conversion. Please try again later.')
