import pandas as pd

# Prompt for the source location and file name
sourceLocation = input("Enter source file location: ")
sourceFileName = input("Enter source file name: ")

# Read the source workbook into a pandas DataFrame
source_df = pd.read_excel(f"{sourceLocation}/{sourceFileName}.xlsx")
sourceSheetName = input("Enter source sheet name: ")

# Set the ranges for data in the source DataFrame
sourceData = source_df[['First Name', 'Last Name', 'Company', 'Email', 'Country',
                         'Mobile Phone', 'Keywords', 'Website', 'Company Linkedin Url', 'Person Linkedin Url', 'Industry']]

# Create a new target DataFrame
target_df = pd.DataFrame(columns=[
    "First Name", "Last Name", "Account Name", "HQ Sales", "Assigned User Name",
    "Product Interests", "Country", "Division", "Email Address", "Mobile Phone",
    "Description", "Title", "Website", "Lead Source", "Person Linkedin Url",
    "WeChat", "Company Linkedin Url", "Industry", "Who Initiated the first contact?",
    "Department", "Reports To", "Facebook Account", "Teams"
])

# Copy data from source to target
target_df["First Name"] = sourceData["First Name"]
target_df["Last Name"] = sourceData["Last Name"]
target_df["Account Name"] = sourceData["Company"]
target_df["Country"] = sourceData["Country"]
target_df["Email Address"] = sourceData["Email"]
target_df["Mobile Phone"] = sourceData["Mobile Phone"]
target_df["Description"] = sourceData["Keywords"]
target_df["Website"] = sourceData["Website"]
target_df["Person Linkedin Url"] = sourceData["Person Linkedin Url"]
target_df["Company Linkedin Url"] = sourceData["Company Linkedin Url"]
target_df["Industry"] = sourceData["Industry"]

# Fill columns with permanent info
product_interests = "Fun Walls, Ropes Course, Rollglider, Adventure Trail, Caving, Cloud Climb, Zip Line, Ninja Course, Tree Course, Slides"
division = "Active Entertainment"
lead_source = "LinkedIn"
first_contact = "I first contacted client"
teams = "Global"

target_df["Product Interests"] = product_interests
target_df["Division"] = division
target_df["Lead Source"] = lead_source
target_df["Who Initiated the first contact?"] = first_contact
target_df["Teams"] = teams

# Fill specified columns in the target DataFrame
HQSales = input("Who is the head sales: ")
assignedTo = input("To whom is this lead assigned to: ")
target_df["HQ Sales"] = HQSales
target_df["Assigned User Name"] = assignedTo

# Delete rows in the target DataFrame based on conditions
target_df = target_df[target_df["Email Address"].str.contains('@', na=False)]
target_df["Last Name"].fillna(target_df["First Name"], inplace=True)

# Replace country names
country_replacements = {
    "United States": "USA",
    "South Korea": "Korea, South",
    "Republic of Indonesia": "Indonesia",
    # Add other country replacements as needed
}
target_df["Country"].replace(country_replacements, inplace=True)

# Replace industry names
industry_replacements = {
    "real estate": "Real Estate Developer",
    "architecture & planning": "Design",
    "civil engineering": "Engineering",
    # Add other industry replacements as needed
}
industry_list = ["Real Estate Developer", "Design", "Engineering", "Family Entertainment Center", "Seaside Resort",
                 "Airport", "Family Resort", "Consulting", "Time shared resort"]

target_df["Industry"].replace(industry_replacements, inplace=True)
target_df["Industry"].fillna("", inplace=True)

# Prompt for the target location and file name
targetLocation = input("Where should I save the file: ")
targetFileName = input("Pick a name: ")

# Save the target DataFrame to Excel
target_df.to_excel(f"{targetLocation}/{targetFileName}.xlsx", index=False)
