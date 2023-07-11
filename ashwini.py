import docx
import json
def convert_word_to_json(file_path):
    doc = docx.Document(file_path)
#{
#        "Category": "Education, Houses, Landed Property,Conveyances, Mother, General Happiness",
#        "House": 4,
#        "Planet": "Moon",
#		"Sign" : "Libra",
#        "Nakshatra": "Chitra",
#        "Auspicious Keywords": "Care more for peace and pleasure and harmony,good in education, robust health; pleasant mother,  happy domestic life,  change of residence,  gains money through land, mine"
#    },
 
    nakshatra = ''
    house = ''
    category = ''
    auspicious = ''
    inauspicious = ''
    is_auspicious = True
    data = []
    is_first = True

    for paragraph in doc.paragraphs:
        style = paragraph.style.name
        if style == "Heading 1":
            #print("Heading 1", paragraph.text.strip())
            nakshatra= paragraph.text.strip()

        elif style == "Heading 2":
            #print("is_first", is_first)
            if is_first is False:
                obj = {
                    "Category":category,
                    "House": house,
                    "Planet": "Moon",
                    "Nakshatra": nakshatra,
                    "Auspicious": auspicious,
                    "Inauspicious": inauspicious
                }
                data.append(obj)
            auspicious = ''
            inauspicious = ''
            #print("Heading 2", paragraph.text.strip()[0])
            house = paragraph.text.strip()[:2]
        elif style == "Normal":
            #print("Heading 1", paragraph.text.strip())
            category= paragraph.text.strip()
        elif style == "Heading 3":
            is_first = False
            #print("Heading 3", paragraph.text.strip()[0])
            if paragraph.text.strip()[0] == 'A':
                is_auspicious = True
            else:
                is_auspicious = False
        elif style == "Normal (Web)":
            #print("Normal", paragraph.text.strip())
            if is_auspicious:
                auspicious += paragraph.text.strip() + " "
            else:
                inauspicious += paragraph.text.strip() + " "
    obj = {
        "Category":category,
        "House": house,
        "Planet": "Moon",
        "Nakshatra": nakshatra,
        "Auspicious": auspicious,
        "Inauspicious": inauspicious
    }
    data.append(obj)
    json_data = json.dumps(data, indent=4)
    return json_data
word_file_path = 'Ashwini1.docx'
json_data = convert_word_to_json(word_file_path)
#print(json_data)
# Write the JSON data to a file
output_file_path = "ashwini.json"
with open(output_file_path, 'w') as output_file:
    output_file.write(json_data)
print("Conversion complete. JSON data written to", output_file_path)