
from pickle import TRUE
import openpyxl
import os
action_arr=[]
expected_arr=[]
random_arr=[]
testcase=''
action_text=''
xml1=''
xml_final_text=''
files_name=[]
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    fname=f.endswith('.xlsx')
    # print(fname)
    if(fname):
        files_name.append(f)
 
# print(files_name)
excel_file_name=files_name[0]
wb_obj = openpyxl.load_workbook(excel_file_name) 

sheet = wb_obj.active
# print(sheet['G2'].value)
maxnum=0
no_of_testcase=0
for i in range(100):
    
    txt=''
    txt=str(sheet[f'B{i+1}'].value)
    # print(txt)
    if(len(txt)>4):
        maxnum=maxnum+1
    else:
        break
# print("maxnum is:::::::::",maxnum)



flag=False
arr=[]
# print("maxnum is",maxnum)
for row in sheet.iter_rows(max_row=maxnum):
        no_of_testcase+=1
        for cell in row:
            if(flag):
                arr.append(cell.value)        
        action_text=''
        if(flag): 
            Test_Objective=arr[0]
            Module=arr[1]
            User_Story=arr[2]
            User_Story=User_Story.strip()
            
            # print(Test_Objective)
            # print(Module)
            # print(User_Story)
            AC_Mapping=arr[3]
            Importance=arr[4]
            print('importance is:',Importance)
            Importance=str(Importance.strip())
            Importance=Importance.lower()
            if(Importance=="medium"):
                Importance=2
            elif(Importance=="low"):
                Importance=1
            elif(Importance=="high"):
                Importance=3
            else:
                print("something wrong with Importance")
            # print(Importance)
            Pre_Requisite=arr[5]
            precondition_text=''
            prequisite_flag=False
            for i in range(1,100):
                if(f"{i}." in Pre_Requisite):
                    pre_index=Pre_Requisite.find(f"{i}.")
                    pre_index2=Pre_Requisite.find(f"{i+1}.")
                    if(pre_index2==-1):
                        pre_text1=Pre_Requisite[pre_index:]
                        prequisite_flag=True

                        # test_steps=test_steps.replace(f'Step {k}',"")
                    else:
                        pre_text1=Pre_Requisite[pre_index:pre_index2]
                        # print("text is:::::",text1)
                    pre_text1=pre_text1.strip()
                    if(prequisite_flag):
                        precondition_text=precondition_text+pre_text1
                    else:
                        precondition_text=precondition_text+pre_text1+f"<br/>"
                else:
                    precondition_text=precondition_text+Pre_Requisite
                    prequisite_flag=True


                if(prequisite_flag):
                    break
                


            Actions=arr[6]

            Expected_Result=arr[7]

            for l in range(1,100):
                if(f"{l}." in Actions):
                    random_index=Actions.find(f"{l}.")

                    random_index2=Actions.find(f"{l+1}.")
                    if(random_index2==-1):
                        text1=Actions[random_index:]
                        
                            # test_steps=test_steps.replace(f'Step {k}',"")
                    else:
                        text1=Actions[random_index:random_index2]
                        # print("text is:::::",text1)
                    text1=text1.strip()
                    random_arr.append(text1)#storing all the Initial action steps in
            for x in range(1,100):
                if(f"{x}." in Expected_Result):
                    index = Expected_Result.find(f"{x}.")
                    index2=Expected_Result.find(f"{x+1}.")
                    if(index2==-1):
                        text1=Expected_Result[index:]
                        flag2=True
                    else:
                        text1=Expected_Result[index:index2]
                    text1=text1.strip()
                    expected_arr.append(text1)

            xml2=''
            xml2=xml2+f''' 
    <testcase internalid="" name="{User_Story}">
        <node_order><![CDATA[]]></node_order>
        <externalid><![CDATA[]]></externalid>
        <version><![CDATA[]]></version>
        <summary><![CDATA[{Test_Objective}]]></summary>
        <preconditions><![CDATA[{precondition_text}]]></preconditions>
        <execution_type><![CDATA[1]]></execution_type>
        <importance><![CDATA[{Importance}]]></importance>
        <steps>'''
            t=1
            # print('length of random array is:',len(random_arr))
            for i in range(len(random_arr)):
                try:
                    a=random_arr[i]
                except:
                    print("/nThere is some mistake in action step")
                    break
                try:
                    b=expected_arr[i]  
                except:
                    print("/nThere is some mistake in expected result step")
                    break
                a=a.replace('‘',"'")
                a=a.replace('’',"'")
                a=a.replace('’',"'")
                a=a.replace('”','"')
                a=a.replace('“','"')
                b=b.replace('‘',"'")
                b=b.replace('’',"'")
                b=b.replace('’',"'")
                b=b.replace('”','"')
                b=b.replace('“','"')
                if(i>8):
                    a=a.strip()
                    b=b.strip()
                    text_b=b[0:3]
                    b=b.replace(text_b,"")
                    b=b.strip()
                    text_b=a[0:3]
                    a=a.replace(text_b,"")
                    a=a.strip()
                else:
                    a=a.strip()
                    b=b.strip()
                    text_b=b[0:2]
                    b=b.replace(text_b,"")
                    b=b.strip()
                    text_b=a[0:2]
                    a=a.replace(text_b,"")
                    a=a.strip()
                
                    
               
                action_text=action_text+f'''
            <step>

                    <step_number><![CDATA[{t}]]></step_number>
                    <actions><![CDATA[{a}]]>
                    </actions>

                        <expectedresults><![CDATA[{b}]]>
                        </expectedresults>
                        <execution_type><![CDATA[1]]></execution_type>
                
                    </step>'''
                t=t+1
                    
            # print('action text is::',action_text)

            xml_final_text=f'''    </steps>

        

            </testcase>

            '''
            testcase=testcase+xml2+action_text+xml_final_text

        arr=[]
        action_arr=[]
        expected_arr=[]
        random_arr=[]
        flag=True 

start_text=f'''<?xml version="1.0" encoding="UTF-8"?>

<testcases>

'''

end_text=f'''



</testcases>'''


xml_file_text=start_text+testcase+end_text
# print(xml_file_text)
with open('readme.xml', 'w') as f:
    f.write(xml_file_text)

print("XML CREATED SUCCESSFULLY")
print("Number Of Testcases is:",no_of_testcase-1)
input()

# For expected Result you can add the data by removing the number and extra spaces and make changes accordance in last this 
# will be more benefecial