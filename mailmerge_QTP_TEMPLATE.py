from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

# Save the TEMPLATE document in "template" variable
template = "QTP-TEMPLATE_try_mailmerge_python.docx"

# Save the template as a MailMerge object. Note: describe this better, not sure if accurately described
document = MailMerge(template)

# Print out to console, the get_merge_fields 
# print(document.get_merge_fields())

print("You're merge fields are:\n")
for x in document.get_merge_fields():
    print(x)
print("\n")

document.merge(
    DO160_S4_Cat='A4',
    DO160_S7_Cat='B',
    Product_Description='75 INCH MONITOR WITH AVB AUDIO',
    Part_Number='U750A103',
)

document.write('text-output.docx')