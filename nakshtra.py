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
    sign = ''
    auspicious = ''
    inauspicious = ''
    is_auspicious = True
    data = []
    is_first = True
    is_next =True

    for paragraph in doc.paragraphs:
        style = paragraph.style.name
        if style == "Heading 1":
            
            nakshatra= paragraph.text.strip()
            if nakshatra == 'Pushya':
                sign = "Cancer"
            else:
                sign = "Libra"

        elif style == "Heading 2":
            if is_first is False:
                obj = {
                    "Category":category,
                    "House": house,
                    "Planet": "Moon",
                    "Sign": sign,
                    "Nakshatra": nakshatra,
                    "Auspicious": auspicious,
                    "Inauspicious": inauspicious
                }
                data.append(obj)
            auspicious = ''
            inauspicious = ''
            house = paragraph.text.strip()[:2]
        elif style == "Normal":
            category= paragraph.text.strip()
        elif style == "Heading 3":
            is_first = False
            #is_next = False      
            if paragraph.text.strip()[0] == 'A':
                is_auspicious = True
            else:
                is_auspicious = False
        elif style == "Heading 6":  
            #is_next = False      
            if is_auspicious:
                auspicious += paragraph.text.strip() + " "
            else:
                inauspicious += paragraph.text.strip() + " "
    obj = {
        "Category":category,
        "House": house,
        "Planet": "Moon",
        "Sign": sign,
        "Nakshatra": nakshatra,
        "Auspicious": auspicious,
        "Inauspicious": inauspicious
    }
    data.append(obj)
    json_data = json.dumps(data, indent=4)
    return json_data
word_file_path = 'nakshtra.docx'
json_data = convert_word_to_json(word_file_path)

output_file_path = "nakshtra.json"
with open(output_file_path, 'w') as output_file:
    output_file.write(json_data)
print("Conversion complete. JSON data written to", output_file_path)