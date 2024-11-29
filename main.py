from docx import Document
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os

file_template_data = os.getcwd() + "/template info - SHORT - Individual.txt"

def find_duplicate(lst):
    seen = set()
    for element in lst:
        if element in seen:
            return element.strip()
        seen.add(element)
    return None  # If no duplicates are found


def get_first_names(names):
    first_names = []
    for name in names.split(","):
        first_names.append(name.strip().replace("and ", "").split(" ")[0].strip())
    return first_names

def find_element_index(last_names, duplicate_lastname):
    indexes_found = []
    for i, last_name in enumerate(last_names):
        if last_name == duplicate_lastname:
            indexes_found.append(i)  # Save the index of matching last name
    return indexes_found


def add_names_to_duplicate_lastnames(client_names, mr_mrs_client_last_name):
    last_names = [i.strip() for i in mr_mrs_client_last_name.replace(", and ", ",").split(",")]
    duplicate_lastname = find_duplicate(last_names)

    first_names = get_first_names(client_names)

    for index in find_element_index(last_names, duplicate_lastname):
        last_names[index] = f"{last_names[index]} ({first_names[index]})"

    last_names[-1] = "and " + last_names[-1]
    return ", ".join(last_names)

def custom_title(text, excluded_words=None):
    """
    Capitalizes the first letter of each word in a string,
    except for 'and' (case-sensitive) and words in the excluded_words list.

    :param text: The input string.
    :param excluded_words: A list of words to exclude from capitalization.
    :return: The formatted string.
    """
    if excluded_words is None:
        excluded_words = []

    # Create a set of excluded words for O(1) lookup time
    excluded_words_set = set(excluded_words)

    return " ".join(
        word if word.lower() in excluded_words_set else word.capitalize()
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
IS_YOUNG = "young" if IS_YOUNG == "yes" else (IS_YOUNG if len(IS_YOUNG) > 1 else "")

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
    for index, client_sex in enumerate(CLIENT_SEX.split(",")):
        if client_sex.strip() == "woman":
            if IS_YOUNG.split(",")[index].strip() == "no":
                CLIENT_TITLE.append("Mrs.")
            else:
                CLIENT_TITLE.append("Ms.")

        elif client_sex.strip() == "minor":
            CLIENT_TITLE.append("")
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

if " and " not in CLIENT_NAME:
    CLIENT_LAST_NAME = CLIENT_NAME_EACH_CAP.split(" ")[-1] if CLIENT_NAME_EACH_CAP.split(" ")[-1] not in ["Sr", "Jr"] else CLIENT_NAME_EACH_CAP.split(" ")[-2] + " " +CLIENT_NAME_EACH_CAP.split(" ")[-1]
    MR_MRS_CLIENT_NAME_EACH_CAP = (CLIENT_TITLE + CLIENT_NAME_EACH_CAP).title()
    MR_MRS_CLIENT_NAME_ALL_CAP = (CLIENT_TITLE + CLIENT_NAME_EACH_CAP).upper()
    MR_MRS_CLIENT_LAST_NAME = CLIENT_TITLE + CLIENT_LAST_NAME
    
elif " and " in CLIENT_NAME and "," not in CLIENT_NAME: # more than one client 

    MR_MRS_CLIENT_LAST_NAME = ""
    MR_MRS_CLIENT_NAME = ""

    for index, client_name in enumerate(CLIENT_NAME.split(" and ")):
        client_name = client_name.strip()  # Remove extra spaces

        # Add " and " if there's already a name
        if MR_MRS_CLIENT_LAST_NAME:
            MR_MRS_CLIENT_LAST_NAME += ", and "
        if MR_MRS_CLIENT_NAME:
            MR_MRS_CLIENT_NAME += ", and "

        # Check for "Sr" or "Jr" in the name
        if any(title in client_name for title in ["Sr", "Jr"]):
            # Add full name and last two words for titles
            if CLIENT_TITLE[index]:
                MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
                MR_MRS_CLIENT_LAST_NAME += (
                    CLIENT_TITLE[index] + " " + " ".join(client_name.split()[-2:])
                )
            else:
                MR_MRS_CLIENT_NAME += client_name
                MR_MRS_CLIENT_LAST_NAME += client_name
        else:
            if CLIENT_TITLE[index]:
                # Add full name and last word (last name)
                MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
                MR_MRS_CLIENT_LAST_NAME += CLIENT_TITLE[index] + " " + client_name.split()[-1]
            else:
                MR_MRS_CLIENT_NAME += client_name
                MR_MRS_CLIENT_LAST_NAME += client_name

    MR_MRS_CLIENT_NAME_EACH_CAP = custom_title(MR_MRS_CLIENT_NAME, ["and"])
    MR_MRS_CLIENT_NAME_ALL_CAP = MR_MRS_CLIENT_NAME.upper()

else:
    MR_MRS_CLIENT_LAST_NAME = ""
    MR_MRS_CLIENT_NAME = ""

    for index, client_name in enumerate(CLIENT_NAME.split(", ")):
        last_name_tracker = []

        client_name = client_name.strip("and ").strip()  # Remove extra spaces

        if index != len(CLIENT_NAME.split(", "))-1:
            # Add ", " if there's already a name
            if MR_MRS_CLIENT_LAST_NAME:
                MR_MRS_CLIENT_LAST_NAME += ", "
            if MR_MRS_CLIENT_NAME:
                MR_MRS_CLIENT_NAME += ", "

        else:
            # Add ", " if there's already a name
            if MR_MRS_CLIENT_LAST_NAME:
                MR_MRS_CLIENT_LAST_NAME += ", and "
            if MR_MRS_CLIENT_NAME:
                MR_MRS_CLIENT_NAME += ", and "

    #     # Check for "Sr" or "Jr" in the name
        if any(title in client_name for title in ["Sr", "Jr"]):
            # Add full name and last two words for titles
            if CLIENT_TITLE[index]:
                MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
                MR_MRS_CLIENT_LAST_NAME += (
                    CLIENT_TITLE[index] + " " + " ".join(client_name.split()[-2:])
                )
            else:
                MR_MRS_CLIENT_NAME += client_name
                MR_MRS_CLIENT_LAST_NAME += client_name

        else:
            last_name_tracker.append(client_name.split()[-1])
            if CLIENT_TITLE[index]:
                
                MR_MRS_CLIENT_NAME += CLIENT_TITLE[index] + " " + client_name
                MR_MRS_CLIENT_LAST_NAME += CLIENT_TITLE[index] + " " + client_name.split()[-1]

            else:
                MR_MRS_CLIENT_NAME += client_name
                MR_MRS_CLIENT_LAST_NAME += client_name

    MR_MRS_CLIENT_NAME_EACH_CAP = custom_title(MR_MRS_CLIENT_NAME, ["and"])
    MR_MRS_CLIENT_NAME_ALL_CAP = MR_MRS_CLIENT_NAME.upper()

if " and " in MR_MRS_CLIENT_LAST_NAME:
    MR_MRS_CLIENT_LAST_NAME = add_names_to_duplicate_lastnames(CLIENT_NAME, MR_MRS_CLIENT_LAST_NAME)

MR_MRS_CLIENT_LAST_NAME = MR_MRS_CLIENT_LAST_NAME.title()

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
    elif "minor" in CLIENT_SEX:
        if "woman" in CLIENT_SEX.split(", ") and "man" in CLIENT_SEX.split(", "):
            CLIENT_SEX = f"healthy men, women, or minors lose"
        elif "woman" in CLIENT_SEX.split(", "):
            CLIENT_SEX = f"healthy women, or minors lose"
        else:
            CLIENT_SEX = f"healthy men, or minors lose"
    else:
        CLIENT_SEX = f"healthy men or women lose"
else:
    CLIENT_SEX = f"a healthy {CLIENT_SEX} loses"

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
    "CLAIM_RESPONSIBLE_RECEIVER": CLAIM_RESPONSIBLE_RECEIVER.title(),
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
