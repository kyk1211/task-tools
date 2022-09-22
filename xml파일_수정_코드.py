import os
import xml.etree.ElementTree as ET

## 수정할 파일 위치 디렉토리
currentDir = os.path.dirname(os.path.realpath(__file__))
targetDir = rf'{currentDir}\work'
num = 1

##targetDir에서 .xml파일 이름들 리스트로 가져오기
file_list = os.listdir(targetDir)
xml_list = []
for file in file_list:
    if file.endswith('.xml'):
        xml_list.append(file)

##모든 .xml파일에 대해 수정
for xml_file in xml_list:
    target_path = targetDir + '/' + xml_file
    targetXML = open(target_path, 'rt', encoding='UTF8')

    tree = ET.parse(targetXML)

    root = tree.getroot()
    
    ##수정할 부분
    xml_file = xml_file.split('.')[0]
    target_tag1 = root.find("mgnumber")
    target_tag2 = root.find("filename")
    target_tag1.text = xml_file[:-2]
    target_tag2.text = xml_file + '.jpg'
    print("[" + str(num) + "]" + xml_file + "[success]")
    
    tree.write(target_path, encoding='utf8', xml_declaration=False)
    num += 1

print("finished")