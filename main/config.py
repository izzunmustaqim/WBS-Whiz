prompt_list_task = (
                        "#Input"
                        "{screen_layout_json}"
                        "{app_detailed_spec_data_converted_json}"
 
                        "#INSTRUCTIONS"
                        "As a project manager, create the task details for the project. The inputs are the screen layout details and Application Detailed Specification Files."
                        "Given the image flowchart for reference of the whole structure of the flow for the application"
                        "Please provide the information in table format as below"
                        "Please sort the task considering logic that the screen should have completed first then the function can be developed"
                        "If any of the task is related or can be grouped as one, you may do so"
                        "Please make sure the first row of each task is not the title or header for the subtask. Instead, it should be the initial task before the subtask"
                        "Please state the complexity of each task priority"
                        "It should consist only of the following:"
                        "No       |   Task Description          |    Complexity"
                        "1.       |     Task 1                  |    hard/mediuam/easy"
                        "   1.1   |         Subtask 1           |    hard/medium/easy"
                        "2.       |     Task 2                  |    hard/medium/easy"
                    )

prompt = (
            "#Input"
            "{task_details_data} "
            "{skill_set_data}"

            "#INSTRUCTIONS"
            "As a project manager, create a WBS for the project. The input are the task details and the skill set of the team members."
            "Please give in table format where each of the title should be the header"
            "For high task priority assign first to senior developer then to middle developer"
            "For medium task priority assign first to middle developer then junior developer" 
            "For low task priority first assign to junior developer then middle developer"
            "Please also consider the skills needed for each task with the member with high level skills for that particular need"
            "The output should include all members for each task and one task for one person only"
            "Please estimate the duration of each task without including it in the result. Then, set the start date and end date based on the duration."
            "Please use the date range from the input data for start date = {start_date_str} and end date = {end_date_str} for each of the task."
            "Please fully utilize the date but do not exceed the project's end date at {end_date_str}"
            "It should consists of:"
                
            "| Item No. | Task Description: {task_description} | Assigned to: {assigned_to} | Progress: {progress} | Plan Start date: {plan_start_date} | Plan End date: {plan_end_date} |\n"
            # "| Item No. | Task Description: {Task Description} | Assigned to: {Assigned to based on the input data of the task} | Progress: {Progress of the task either To do/In progress/Waiting for Review/Done, default is To do} | Plan Start date: {Start date based on the input data of the task} | Plan End date: {End date based on the input data of the task} |\n"
        )
error_message = {
    "FileNotFoundError": "The file was not found. Please check the file path and try again.",
    "FolderNotFoundError": "No folder selected. Please choose a folder and try again.",
    "ManyExcelError": "Too many Excel files detected. The folder must contain no more than 5 Excel files. Please check the folder and try again.",
    "EmptyDataError": "The file is empty. Please provide a valid Excel file with data.",
    "ParserError": "There was a problem parsing the file. Please ensure the file is a valid Excel file.",
    "FileTooBig": "The file was to big. Please check the file and try again.",
    "APIKeyError": "The API Key file was not found. Please check the file path and try again.",
    "FailReadError": "Failed to read Excel file: {error_message}",
    "APIEmptyField": "API Key cannot be empty",
    "InvalidKeyError": "Invalid API Key format. Please enter a valid API key format: \n\n48 character long and it contains a mix of uppercase letters, lowercase letters, and digits",
    "FullWidthCharacterError": "API Key contains full-width characters, which are not allowed.",
    "GeneralError": "An error occurred: {error_message}"
}