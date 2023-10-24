import PySimpleGUI as sg
from abn_api import abn
from response import api_response

def main():
    layout = [
        [sg.Text("ABN Lookup Tool", font= 16)],
        [sg.Text("Name:"), sg.InputText(key="name")],
        [sg.Text("Post Code:"), sg.InputText(key="postcode")],
        [sg.Text("Number of Record:"), sg.InputText(key="n")],
        [sg.Button("Search"), sg.Button("Result Detail"), sg.Button("CSV Download"), sg.Button("Exit")],
        [sg.Output(size=(60, 30), key='-OUTPUT-', font= 16)],
    ]

    window = sg.Window("ABN Lookup Tool", layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Search":
            try:
                name = values["name"]
                postcode = values["postcode"]
                n = values["n"]

                if name:
                    abn_instance = abn(name, postcode)
                    response = abn_instance.open()
                    returnedXML = response.read()
                    response_parser = api_response()
                    result = response_parser.parse_response(returnedXML)
                    if n == "":
                        n = 10
                    elif int(n) > 0:
                        n = int(n)
                    elif int(n) > len(result):
                        n = len(result)   
                    message = f"ABN Results of {name} with postcode {postcode}:\n{result.head(n)}\n"
                    window['-OUTPUT-'].update(message)

            except Exception as e:
                print(f"An error occurred: {str(e)}")
                print()        

        elif event == "Result Detail":
            try:
                if 'result' in locals() and result is not None:
                    count_df = len(result)
                    summary_postcode = result["PostCode"].value_counts()
                    message = f"The ABN Results with keyword '{name}' with postcode '{postcode}' have returned {count_df} output record. \n"
                    message2 = f"Top 40 Post code distribution:\n{summary_postcode.head(40).to_string()}"
                    messages = [message, message2]
                    formatted_message = '\n'.join(messages)
                    window['-OUTPUT-'].update(formatted_message)
                else:
                    message = "No search results available."
                    window['-OUTPUT-'].update(message)

            except Exception as e:
                print(f"An error occurred: {str(e)}")

        elif event == "CSV Download":
                    try:
                        if 'result' in locals() and result is not None:
                            # Display a file dialog to select where to save the CSV file
                            csv_file_path = sg.popup_get_file("Save CSV file", save_as=True, file_types=(("CSV Files", "*.csv"),))
                            if csv_file_path:
                                # Save the DataFrame as a CSV file
                                result.to_csv(csv_file_path, index=False)
                                sg.popup(f"DataFrame saved to {csv_file_path}")
                        else:
                            sg.popup("No search results available.")

                    except Exception as e:
                        sg.popup_error(f"An error occurred: {str(e)}")
    window.close()

if __name__ == "__main__":
    main()

# python main.py
