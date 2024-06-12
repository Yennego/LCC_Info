import os
import streamlit as st
import pandas as pd
from PIL import Image

# Define the path to the Excel file and images folder
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_file = os.path.join(current_dir, 'membership_info.xlsx')
images_folder = os.path.join(current_dir, 'images')

# Ensure the images folder exists
os.makedirs(images_folder, exist_ok=True)

# Load the Excel file and ensure certain columns are treated as strings
if os.path.exists(excel_file):
    df = pd.read_excel(excel_file, dtype={
        'ID': str,
        'Contact Number 1': str,
        'Contact Number 2': str,
        'Emergency Contact Cell': str,
        'Joined Year': str,
        'Emergency Contact Name': str
    })
else:
    df = pd.DataFrame(columns=[
        "ID", "Last Name", "First Name", "Middle Name", "Date of Birth", "Place of Birth",
        "Nationality", "Age", "Sex", "Home Address", "Contact Number 1", "Contact Number 2",
        "Position/Rank", "Occupation", "Marital Status", "Spouse Name", "Date of Marriage",
        "Born Again", "Baptized", "Membership Type", "Joined Year", "Email",
        "Emergency Contact Name", "Emergency Contact Cell", "Image Path"
    ])

# Streamlit app
st.title("LCC Membership Form")

# Form to input new data
with st.form(key='membership_form'):
    id_ = st.text_input("ID")
    last_name = st.text_input("Last Name")
    first_name = st.text_input("First Name")
    middle_name = st.text_input("Middle Name")

    dob = st.text_input("Date of Birth (YYYY-MM-DD)")
    place_of_birth = st.text_input("Place of Birth")

    nationality = st.text_input("Nationality")
    age = st.number_input("Age", min_value=0, max_value=120)
    sex = st.radio("Sex", ('Male', 'Female'))

    home_address = st.text_input("Home Address")
    contact_number1 = st.text_input("Contact Number 1")
    contact_number2 = st.text_input("Contact Number 2")

    position_rank = st.text_input("Position/Rank in LCC")
    occupation = st.text_input("Occupation")

    marital_status = st.radio("Marital Status", ['Single', 'Married', 'Divorced', 'Separated'])
    spouse_name = st.text_input("Name of Spouse (if married)")
    date_of_marriage = st.text_input("Date of Marriage (YYYY-MM-DD, if married)")

    born_again = st.radio("Are you born again?", ['Yes', 'No'])
    baptized = st.radio("Are you baptized according to Matthew 28:19?", ['Yes', 'No'])

    membership_type = st.radio("Type of Membership", ['Full Membership', 'Associate Membership'])
    joined_year = st.text_input("What year did you join the LCC family?")
    email = st.text_input("Email")

    emergency_contact_name = st.text_input("Emergency Contact Name")
    emergency_contact_cell = st.text_input("Emergency Contact Cell")

    uploaded_file = st.file_uploader("Upload Image", type=['jpg', 'png', 'jpeg'])

    if st.form_submit_button("Submit"):
        if id_ and last_name and first_name and dob and place_of_birth and contact_number1 and email:
            # Save the image
            image_path = ""
            if uploaded_file:
                image_path = os.path.join(images_folder, uploaded_file.name)
                with open(image_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

            # Add the new data to the DataFrame
            new_data = pd.DataFrame({
                "ID": [id_], "Last Name": [last_name], "First Name": [first_name], "Middle Name": [middle_name],
                "Date of Birth": [dob], "Place of Birth": [place_of_birth],
                "Nationality": [nationality], "Age": [age], "Sex": [sex],
                "Home Address": [home_address], "Contact Number 1": [contact_number1], "Contact Number 2": [contact_number2],
                "Position/Rank": [position_rank], "Occupation": [occupation],
                "Marital Status": [marital_status], "Spouse Name": [spouse_name], "Date of Marriage": [date_of_marriage],
                "Born Again": [born_again], "Baptized": [baptized],
                "Membership Type": [membership_type], "Joined Year": [joined_year], "Email": [email],
                "Emergency Contact Name": [emergency_contact_name], "Emergency Contact Cell": [emergency_contact_cell],
                "Image Path": [image_path]
            })
            df = pd.concat([df, new_data], ignore_index=True)

            # Ensure certain columns remain strings
            df['ID'] = df['ID'].astype(str)
            df['Contact Number 1'] = df['Contact Number 1'].astype(str)
            df['Contact Number 2'] = df['Contact Number 2'].astype(str)
            df['Emergency Contact Cell'] = df['Emergency Contact Cell'].astype(str)
            df['Joined Year'] = df['Joined Year'].astype(str)
            df['Emergency Contact Name'] = df['Emergency Contact Name'].astype(str)

            # Save the DataFrame to the Excel file
            try:
                df.to_excel(excel_file, index=False)
                st.success("Data submitted successfully!")
            except PermissionError as e:
                st.error(f"PermissionError: {e}")

        else:
            st.error("Please fill in all required fields.")

# Display the DataFrame
# Ensure certain columns remain strings before displaying
df['ID'] = df['ID'].astype(str)
df['Contact Number 1'] = df['Contact Number 1'].astype(str)
df['Contact Number 2'] = df['Contact Number 2'].astype(str)
df['Emergency Contact Cell'] = df['Emergency Contact Cell'].astype(str)
df['Joined Year'] = df['Joined Year'].astype(str)
df['Emergency Contact Name'] = df['Emergency Contact Name'].astype(str)

st.write("### Membership Information")
st.dataframe(df)

# Option to display images
if st.checkbox("Show Images"):
    for index, row in df.iterrows():
        st.write(f"Name: {row['First Name']} {row['Last Name']}, Age: {row['Age']}, Email: {row['Email']}")
        if isinstance(row['Image Path'], str) and os.path.exists(row['Image Path']):
            image = Image.open(row['Image Path'])
            st.image(image, caption=f"{row['First Name']} {row['Last Name']}", use_column_width=True)
        else:
            st.write("No valid image available.")
