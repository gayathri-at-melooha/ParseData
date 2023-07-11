import docx
import json
def convert_word_to_json(file_path):
    doc = docx.Document(file_path)
#{
# 
#        "House": 4,
#        "Planet": "Moon",
#        Keywords": "Care more for peace and pleasure and harmony,good in education, robust health; pleasant mother,  happy domestic life,  change of residence,  gains money through land, mine"
#    },
 
    planet =''
    house = ''
    keyword = ''

    data = []
    is_first = True
    is_planet = True
    for paragraph in doc.paragraphs:
        style = paragraph.style.name
        if style == "Heading 1":
            #is_planet = False   
            planet = paragraph.text.strip()
            
        elif style == "Heading 3":
            
            if is_first == False and is_planet == False:
                obj = {
                    "Planet": planet,
                    "House": house,
                    "Keyword": keyword,
                }
                data.append(obj)
            keyword = ''
            house = paragraph.text.strip()
            #print(house)
            
        elif style == "Normal":
            is_first = False    
            is_planet = False
            keyword += paragraph.text.strip() + " " 
            #print("Normal", paragraph.text.strip())
    obj = {
        "Planet": planet,
        "House": house,
        "Keyword": keyword,
    }
    data.append(obj)

    json_data = json.dumps(data, indent=4)
    return json_data
word_file_path = 'planet_in_house.docx'
json_data = convert_word_to_json(word_file_path)

output_file_path = "planet_in_house.json"
with open(output_file_path, 'w') as output_file:
    output_file.write(json_data)
print("Conversion complete. JSON data written to", output_file_path)