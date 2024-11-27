from docx import Document
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

file_template_data = os.getcwd() + "/template info - SHORT - Individual.txt"

def custom_title(text, excluded_words=None):
    """
    Capitalizes the first letter of each word in a string,
    except for words in the excluded_words list.

    :param text: The input string.
    :param excluded_words: A list of words to exclude from capitalization.
    :return: The formatted string.
    """
    if excluded_words is None:
        excluded_words = []
    
    return " ".join(
        word if word in excluded_words else word.capitalize()
        for word in text.split()
    )


def parse_file_data():
    with open(file_template_data, "r", encoding="utf-8") as file:
        return [i.strip().split(":")[1].strip() for i in file.readlines()]

DATA = parse_file_data()

# Client Information
CLIENT_NAME = DATA[0]
CLIENT_SEX = DATA[1]
IS_YOUNG = DATA[2]
IS_YOUNG = "young" if IS_YOUNG == "yes" else ""

# Insured Information
INSURED_NAME = DATA[3]
INSURED_SEX = DATA[4]
INSURED_TITLE = "Mr. " if INSURED_SEX.lower() == "man" else "Mrs. "

# Contact and Claim Information
VIA_TYPE = DATA[5]
INSURANCE_NAME = DATA[6]
CLAIM_NUMBER = DATA[7]
DATE_OF_LOSS = DATA[8]
CLAIM_RESPONSIBLE_RECEIVER = DATA[9]

# California Civil Code Text
CALIFORNIA_CVC_TEXT = DATA[10]

INSURANCE_INIT = INSURANCE_NAME.split(" ")[0]
INSURANCE_NAME_CAP = INSURANCE_NAME.upper()

# Automatically set gender-specific variables based on CLIENT_SEX
if CLIENT_SEX == "woman":
    HE_SHE_CLIENT = "she"
    HER_HIM_CLIENT = "her"
    HER_HIS_CLIENT = "her"
    HER_HIS_CLIENT_CAP = "HER"
    CLIENT_TITLE = "Ms. " if IS_YOUNG else "Mrs. "
    HE_SHE_CLIENT_PAGE7 = HE_SHE_CLIENT + " was" 

elif CLIENT_SEX == "man":  # CLIENT_SEX == "man"
    HE_SHE_CLIENT = "he"
    HER_HIM_CLIENT = "him"
    HER_HIS_CLIENT = "his"
    HER_HIS_CLIENT_CAP = "HIS"
    CLIENT_TITLE = "Mr. "
    HE_SHE_CLIENT_PAGE7 = HE_SHE_CLIENT + " was" 

else:
    CLIENT_TITLE = []
    for client_sex in CLIENT_SEX.split(","):
        if client_sex.strip() == "woman":
            CLIENT_TITLE.append("Mrs.")
        else:
            CLIENT_TITLE.append("Mr.")

    HE_SHE_CLIENT = "they"
    HER_HIM_CLIENT = "them"
    HER_HIS_CLIENT = "their"
    HER_HIS_CLIENT_CAP = "THEIR"

    # Custom
    HE_SHE_CLIENT_PAGE7 = HE_SHE_CLIENT + " were" 
    
# Automatically set gender-specific variables based on INSURED_SEX
if INSURED_SEX == "woman":
    HE_SHE_INSURED = "she"
    HER_HIM_INSURED = "her"
    HER_HIS_INSURED = "her"

elif INSURED_SEX == "man":
    HE_SHE_INSURED = "he"
    HER_HIM_INSURED = "him"
    HER_HIS_INSURED = "his"


# Format the date as MM/DD/YYYY
SETTLEMENT_EXP_DATE = (datetime.now() + relativedelta(months=1)).strftime("%m/%d/%Y")
SETTLEMENT_EXP_DATE = datetime.strptime(SETTLEMENT_EXP_DATE, "%m/%d/%Y").strftime("%B %d, %Y").upper()

CLIENT_NAME_ALL_CAP = CLIENT_NAME.upper()
CLIENT_NAME_EACH_CAP = custom_title(CLIENT_NAME, ["and"])

if "and" not in CLIENT_NAME:
    CLIENT_LAST_NAME = CLIENT_NAME_EACH_CAP.split(" ")[-1] if CLIENT_NAME_EACH_CAP.split(" ")[-1] not in ["Sr", "Jr"] else CLIENT_NAME_EACH_CAP.split(" ")[-2] + " " +CLIENT_NAME_EACH_CAP.split(" ")[-1]
    MR_MRS_CLIENT_NAME_EACH_CAP = (CLIENT_TITLE + CLIENT_NAME_EACH_CAP).title()
    MR_MRS_CLIENT_NAME_ALL_CAP = (CLIENT_TITLE + CLIENT_NAME_EACH_CAP).upper()
    MR_MRS_CLIENT_LAST_NAME = CLIENT_TITLE + CLIENT_LAST_NAME
    
else: # more than one client 

    MR_MRS_CLIENT_LAST_NAME = ""
    MR_MRS_CLIENT_NAME = ""

    for index, client_name in enumerate(CLIENT_NAME.split("and")):
        client_name = client_name.strip()  # Remove extra spaces

        # Add " and " if there's already a name
        if MR_MRS_CLIENT_LAST_NAME:
            MR_MRS_CLIENT_LAST_NAME += " and "
        if MR_MRS_CLIENT_NAME:
            MR_MRS_CLIENT_NAME += " and "

        # Check for "Sr" or "Jr" in the name
        if any(title in client_name for title in ["Sr", "Jr"]):
            # Add full name and last two words for titles
            MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
            MR_MRS_CLIENT_LAST_NAME += (
                CLIENT_TITLE[index] + " " + " ".join(client_name.split()[-2:])
            )
        else:
            # Add full name and last word (last name)
            MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
            MR_MRS_CLIENT_LAST_NAME += CLIENT_TITLE[index] + " " + client_name.split()[-1]

    MR_MRS_CLIENT_NAME_EACH_CAP = custom_title(MR_MRS_CLIENT_NAME, ["and"])
    MR_MRS_CLIENT_NAME_ALL_CAP = MR_MRS_CLIENT_NAME.upper()

    
INSURED_NAME_ALL_CAP = INSURED_NAME.upper()
INSURED_NAME_EACH_CAP = INSURED_NAME.title()

MR_MRS_INSURED_NAME_EACH_CAP = (INSURED_TITLE + INSURED_NAME_ALL_CAP).title()
MR_OR_MRS_INSURED_NAME_ALL_CAP = MR_MRS_INSURED_NAME_EACH_CAP.upper()
DATE_OF_LOSS_FORMATTED = datetime.strptime(DATE_OF_LOSS, "%m/%d/%Y").strftime("%B %d, %Y")

if "," in CLIENT_SEX:
    if "man, man" == CLIENT_SEX:
        CLIENT_SEX = f"healthy men lose"
    elif "woman, woman" == CLIENT_SEX:
        CLIENT_SEX = f"healthy women lose"
    else:
        CLIENT_SEX = f"healthy men or women loses"
else:
    CLIENT_SEX = f"a healthy {CLIENT_SEX}"



# Store variables in a dictionary
CLIENT_DATA = {
    "CLIENT_NAME": CLIENT_NAME,
    "CLIENT_SEX": CLIENT_SEX,
    "IS_YOUNG": IS_YOUNG,
    "INSURED_NAME": INSURED_NAME,
    "INSURED_SEX": INSURED_SEX,
    "INSURED_TITLE": INSURED_TITLE,
    "HER_HIS_INSURED": HER_HIS_INSURED,
    "HE_SHE_INSURED": HE_SHE_INSURED,
    "VIA_TYPE": VIA_TYPE,
    "INSURANCE_NAME": INSURANCE_NAME,
    "CLAIM_NUMBER": CLAIM_NUMBER,
    "DATE_OF_LOSS": DATE_OF_LOSS,
    "CLAIM_RESPONSIBLE_RECEIVER": CLAIM_RESPONSIBLE_RECEIVER,
    "CALIFORNIA_CVC_TEXT": CALIFORNIA_CVC_TEXT,
    "INSURANCE_INIT": INSURANCE_INIT,
    "HE_SHE_CLIENT": HE_SHE_CLIENT,
    "HE_SHE_CLIENT_PAGE7": HE_SHE_CLIENT_PAGE7,
    "HER_HIM_CLIENT": HER_HIM_CLIENT,
    "HER_HIS_CLIENT": HER_HIS_CLIENT,
    "HER_HIS_CLIENT_CAP": HER_HIS_CLIENT_CAP,
    "MR_MRS_CLIENT_NAME_EACH_CAP": MR_MRS_CLIENT_NAME_EACH_CAP,
    "MR_MRS_CLIENT_NAME_ALL_CAP": MR_MRS_CLIENT_NAME_ALL_CAP,
    "SETTLEMENT_EXP_DATE": SETTLEMENT_EXP_DATE,
    "CLIENT_NAME_ALL_CAP": CLIENT_NAME_ALL_CAP,
    "CLIENT_NAME_EACH_CAP": CLIENT_NAME_EACH_CAP,
    "MR_MRS_CLIENT_LAST_NAME": MR_MRS_CLIENT_LAST_NAME,
    "INSURED_NAME_ALL_CAP": INSURED_NAME_ALL_CAP,
    "INSURED_NAME_EACH_CAP": INSURED_NAME_EACH_CAP,
    "MR_MRS_INSURED_NAME_EACH_CAP": MR_MRS_INSURED_NAME_EACH_CAP,
    "MR_OR_MRS_INSURED_NAME_ALL_CAP": MR_OR_MRS_INSURED_NAME_ALL_CAP,
    "DATE_OF_LOSS_FORMATTED": DATE_OF_LOSS_FORMATTED,
    "INSURANCE_NAME_CAP": INSURANCE_NAME_CAP
}


def edit_docx_preserve_format(doc):
    """
    Replaces target text with replacement text in the document while preserving formatting.
    """
    try:
        # Edit document content
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:  # Access individual formatted text segments (runs)
                if CLIENT_DATA.get(run.text) is not None:
                    run.text = run.text.replace(run.text, CLIENT_DATA[run.text])

        # Edit headers
        for section in doc.sections:
            header = section.header  # Access the header for the section
            for paragraph in header.paragraphs:
                for run in paragraph.runs:
                    if CLIENT_DATA.get(run.text, None):
                        run.text = run.text.replace(run.text, CLIENT_DATA[run.text])

    except Exception as e:
        print(f"An error occurred: {e}")


# Paths for the input and output files
file_path = f"{os.getcwd()}/Template - SHORT - Individual.docx"
output_path = file_path.replace("Template - SHORT - Individual", CLIENT_NAME_ALL_CAP + " - "  + DATE_OF_LOSS_FORMATTED.upper())

# Load the document
doc = Document(file_path)

# Replace placeholders
edit_docx_preserve_format(doc)

# Save the updated document
doc.save(output_path)
print(f"Document updated and saved as: {output_path}")

















def edit_bullet_points(doc, section_title, updated_bullets):
    found_section = False
    bullet_index = 0  # Index for the updated_bullets list

    for paragraph in doc.paragraphs:
        # Locate the section title
        if paragraph.text.strip() == section_title:
            found_section = True
            continue  # Skip to the next paragraph after the section title

        if found_section:
            if paragraph.style.name == "List Paragraph":
                # Update the bullet point text while preserving runs (font/style)
                if bullet_index < len(updated_bullets):
                    runs = paragraph.runs
                    for i, run in enumerate(runs):
                        if i == 0:  # Only update the first run to preserve bullet formatting
                            run.text = updated_bullets[bullet_index]
                        else:
                            run.text = ""
                    bullet_index += 1  # Move to the next bullet point
                else:
                    # Clear any remaining bullet points after all updates
                    for run in paragraph.runs:
                        run.text = ""

            # Stop processing if we've reached any of the next section titles
            if paragraph.text.strip() in ["Diagnosis", "Procedures", "Future"]:
                return  # Exit function once the section ends




# Replace the "Medicine" section with new bullet points
# new_bullets_diagnosis = [
#     "Headache  1 - Migraine Relief",
#     "Headache 2 - Stress Reduction",
#     "Headache 3 - Neurological Treatment"
# ]
# edit_bullet_points(doc, "Diagnosis", new_bullets_diagnosis)



# # Replace the "Medicine" section with new bullet points
# new_bullets_medicine = [
#     "Pill 1 - Migraine Relief",
#     "Pill 2 - Stress Reduction",
#     "Pill 3 - Neurological Treatment"
# ]
# edit_bullet_points(doc, "Procedures", new_bullets_medicine)

# # Replace the "Futures" section with new bullet points
# new_bullets_futures = [
#     "Next Treatment 1 - Follow-up Checkup",
#     "Next Treatment 2 - Advanced Testing",
#     "Next Treatment 3 - Therapy Session"
# ]
# edit_bullet_points(doc, "Future", new_bullets_futures)


# Client data
# CLIENT_NAME = "nataly l sanchez"
# CLIENT_SEX = "woman"  # Change to "man" as needed
# IS_YOUNG = "young"

# # Insured information
# INSURED_NAME = "sabrina r ramirez"
# INSURED_SEX = "woman"  # Change to "man" as needed

# INSURED_TITLE = "Mr. " if INSURED_SEX == "man" else "Mrs. "
# VIA_TYPE = "Email: test@mail.com"
# INSURANCE_NAME = "Geico Insurance Company"
# CLAIM_NUMBER = "CLM123456"
# DATE_OF_LOSS = "01/15/2024"
# CLAIM_RESPONSIBLE_RECEIVER = "Miguel Gonzalez"
# CALIFORNIA_CVC_TEXT = "California Civil Code Section 1542. Shall not trespass."