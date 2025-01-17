prompt_list_task = (
                        "#Input"
                        "{screen_layout_json}"
                        "{app_detailed_spec_data_converted_json}"
                        "{tasks_list_json}"

                        "#INSTRUCTIONS"
                        "As a project manager, determine the complexity for each of the list of task."
                        "The inputs are the screen layout details and Application Detailed Specification Files"
                        "Please assess the complexity of each task described in the following file."
                        "The list of tasks should follow exactly as the given input."
                        "It should consist only of the following:"
                        "No       |   Task Description          |    Complexity"
                        "1.       |     Task 1                  |    Hard/Mediuam/Easy"
                        "2.       |     Task 2                  |    Hard/Mediuam/Easy"
                    )

prompt = (
            "#Input"
            "{task_details_data} "
            "{skill_set_data}"

            "As a project manager, create a WBS for the project. The input are the task details and the skill set of the team members."
            "Please give in table format where each of the title should be the header."
            "In the start of the project, make sure everyone is start at the same date despite the complexity of the task and seniority level."
            "The order for task assignment based on complexity should be following the follow:"
            "Hard complexity task: Senior developer -> Middle developer (only if senior developer not available) -> Junior developer (only if senior and middle developr not available)"
            "Medium complexity task: Middle developer -> Senior developer (only if middle developer not available) -> Junior developer (only if middle and senior developer not available)"
            "Low complexity task: Junior developer -> Middle developer (only if junior developer not available) -> Senior developer (only if junior and middle developer not available)"
            "Please also consider the skills needed for each task with the member with high level skills for that particular need."
            "The task description list should follow exactly like the task details data without the complexity."
            "Please estimate the duration of each task without including it in the result. Then, set the start date and end date based on the duration."
            "Please fully utilize the date, but do not exceed the project's end date at {end_date_str}."
            "Please use the date range from the input data for start date = {start_date_str} and end date = {end_date_str} for each of the task."
            "After assigning all tasks, verify that no developer has unassigned time within the project duration."
            "It should only consist of:"
            "| Item No. | Task Description: {task_description} | Assigned to: {assigned_to} | Progress: {progress} | Plan Start date: {plan_start_date} | Plan End date: {plan_end_date} |\n"
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