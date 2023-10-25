import PySimpleGUI as sg
import pandas as pd
from abn_api import abn
from response import api_response

def main():
    layout = [
        [sg.Text("ABN Lookup Tool", font= 16)],
        [sg.Text("Keywords (comma-separated):"), sg.InputText(key="keywords")],
        [sg.Text("Post Code:"), sg.InputText(key="postcode")],
        [sg.Text("Number of Record:"), sg.InputText(key="n")],
        [sg.Button("Search"), sg.Button("Result Detail"), sg.Button("CSV Download"), sg.Button("Exit")],
        [sg.Output(size=(80, 25), key='-OUTPUT-', font= 16)],
    ]

    window = sg.Window("ABN Lookup Tool", layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Search":
            try:
                # Split the input string into a list of keywords
                keywords = values["keywords"].split(",")
                # Strip leading/trailing whitespace from each keyword
                keywords = [keyword.strip() for keyword in keywords]

                # Initialize an empty DataFrame to store the combined results
                combined_results = pd.DataFrame()

                # Loop through the list of keywords
                for keyword in keywords:
                    name = keyword
                    postcode = values["postcode"]
                    n = values["n"]

                    if name:
                        abn_instance = abn(name, postcode)
                        response = abn_instance.open()
                        returnedXML = response.read()
                        response_parser = api_response()
                        result = response_parser.parse_response(returnedXML)
                        # Append the result to the combined results
                        combined_results = pd.concat([combined_results, result])

                # Drop duplicate records if any
                combined_results = combined_results.drop_duplicates()

                if n == "":
                    n = 10
                elif int(n) > 0:
                    n = int(n)
                elif int(n) > len(combined_results):
                    n = len(combined_results)
                message = f"ABN Results of {keywords} with postcode {postcode}:\n{combined_results.head(n)}\n"
                window['-OUTPUT-'].update(message)

            except Exception as e:
                print(f"An error occurred: {str(e)}")
                print()        

        elif event == "Result Detail":
            try:
                if 'combined_results' in locals() and combined_results is not None:
                    count_df = len(combined_results)
                    summary_postcode = combined_results["PostCode"].value_counts()
                    message = f"The ABN Results with keywords '{keywords}' with postcode '{postcode}' have returned {count_df} output record. \n"
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
                        if 'combined_results' in locals() and combined_results is not None:
                            # Display a file dialog to select where to save the CSV file
                            csv_file_path = sg.popup_get_file("Save CSV file", save_as=True, file_types=(("CSV Files", "*.csv"),))
                            if csv_file_path:
                                # Save the DataFrame as a CSV file
                                combined_results.to_csv(csv_file_path, index=False)
                                sg.popup(f"DataFrame saved to {csv_file_path}")
                        else:
                            sg.popup("No search results available.")

                    except Exception as e:
                        sg.popup_error(f"An error occurred: {str(e)}")
    window.close()

if __name__ == "__main__":
    main()
# python main.py
