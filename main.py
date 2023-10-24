import emlx
import os
import pandas as pd

if __name__ == "__main__":
    print("started")
    folder_path_1 = ("/path")
    folders = [folder_path_1]
    emails = []

    for folder_path in folders:
        for filename in os.listdir(folder_path):
            emlx_file_path = os.path.join(folder_path, filename)

            mail = emlx.read(emlx_file_path)
            body = mail.html
            start_index = body.find("Ik vond de les..:")
            if start_index != -1:
                # Add the length of "Ik vond de les..:" to get the start of the word
                start_index += len("Ik vond de les..:")

                # Find the first word after the start index
                end_index = body.find("<", start_index)

                # Extract the word
                if end_index != -1:
                    score = body[start_index:end_index].strip()
                else:
                    # Handle the case where no word is found
                    print("No word found after 'Ik vond de les..:'")
                    print(body)
            else:
                # Handle the case where the initial substring is not found
                print("Substring 'Ik vond de les..:' not found")
                print(body)

            start_index = body.find("Dit wil ik nog even kwijt..:")
            if start_index != -1:
                # Add the length of "Ik vond de les..:" to get the start of the word
                start_index += len("Dit wil ik nog even kwijt..:")

                # Find the first word after the start index
                end_index = body.find("<", start_index)

                # Extract the word
                if end_index != -1:
                    remark = body[start_index:end_index].strip()
                else:
                    # Handle the case where no word is found
                    print("No word found after 'Dit wil ik nog even kwijt..:'")
                    print(body)
            else:
                # Handle the case where the initial substring is not found
                print("Substring 'Dit wil ik nog even kwijt..:' not found")
                print(body)

            emails.append({'Score': score, 'Remark': remark})
            score = ""
            remark = ""

    df = pd.DataFrame(emails)
    df.to_excel('email_data.xlsx', index=False)
    print("finished")
