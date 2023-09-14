import logging
import sys
import os
import openai
import json
import pandas as pd
import requests

# Set up OpenAI API key
# openai.api_key = os.environ.get('OPENAI_API_KEY')

# Set up OpenAI API key
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT") 
openai.api_version = "2023-05-15"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")

# setup the logger to log to stdout and file
def setup_logger():
    logging.basicConfig(level=logging.DEBUG, filename='matching.log', format='%(asctime)s - %(levelname)s - %(message)s')

    root = logging.getLogger()
    # log >=DEBUG level to file
    root.setLevel(logging.DEBUG)

    handler = logging.StreamHandler(sys.stdout)
    # only log INFO and above to stdout
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(message)s')
    handler.setFormatter(formatter)
    root.addHandler(handler)

# Retrieve mentor/mentee data
def retrieve_data():
    # Read in survey responses CSV files
    responses_df = pd.read_excel("responses.xlsx")

    return responses_df

# Process the data so it can be sent to GPT, ideally into JSON structure.
def preprocess_data(responses_df):
    # Set up authentication headers
    # see https://learn.microsoft.com/en-us/graph/auth/auth-concepts#access-tokens
    access_token = os.environ.get('ACCESS_TOKEN')
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    print(f"Pre-processing data for {len(responses_df)} participants...")

    for index, participant in responses_df.iterrows():
        alias = participant["Email"]

        # Use Microsoft Graph API to get manager
        response = requests.get(f"https://graph.microsoft.com/v1.0/users/{alias}/manager", headers=headers)
        manager = response.json().get('userPrincipalName')
        participant["manager"] = manager
    
        # Use Microsoft Graph API to get skip manager
        response = requests.get(f"https://graph.microsoft.com/v1.0/users/{manager}/manager", headers=headers)
        skip_manager = response.json().get('userPrincipalName')
        participant["skip_manager"] = skip_manager

        # Use Microsoft Graph API to get title
        response = requests.get(f"https://graph.microsoft.com/v1.0/users/{alias}", headers=headers)
        title = response.json().get('jobTitle')
        participant["title"] = title

        #print(f"Found title {title}, manager {manager} and skip manager {skip_manager} for {alias}")

    print(f"Done pre-processing {len(responses_df)} participants")
    return responses_df.to_json()

# Send message to OpenAI API and return the response
def match_with_gpt(inputdata):
    system_message = '''
Help pair mentees with mentors in the provided json data. Each item is either a mentor or mentee. Use the following criteria and constraints provided below:
- Do not use mentors or mentees that aren't provided in the json data
- Do not pair a mentor with more mentees than the mentor has capacity for.
- Mentors and mentees should have a different manager (as specified in the "org" field).
- Mentors should be at least one level above mentees (as specified in the "title" field). 
- Create pairs based on common goals or interests as defined by the objectives and details properties.

Provide a reason/description for why you suggested each pairing as well as a reason why the match might not be ideal.
Provide a rating out of 10 on how close the mentor's and mentee's preferences are aligned. Include whether mentor is over their capacity.

Return the mentor/mentee pairs in the following JSON structure:
{
  "mentor": "mentor email",
  "mentee": mentee email,
  "reason_for": "Reason why Mentor was paired with Mentee",
  "reason_against":"Any reasons why this might not be a good match",
  "alignment_score": "score out of 10",
  "over_capacity":"True if mentor is over their capacity",
}

Let's think step by step and think carefully and logically.
Ensuring no constraints have been ignored, especially the capacity constraint. For mentees who do not have an appropriate match, enter "No match found" in the "mentor" field.

Do not include anything other than valid json in the response. Only include valid matches in the response.
'''
    response = openai.ChatCompletion.create(
        engine = "chat", # use "chat" for GPT-4, "chat35" for GPT-3.5 Turbo
        messages =
            [{"role": "system", "content": system_message},
             {"role": "user", "content": f'{inputdata}'}],
    )
    completion = json.loads(str(response))

    return completion["choices"][0]["message"]["content"]

def postprocess_data(matches):
    # Write matches to Excel
    matches_df = pd.read_json(matches)
    matches_df.to_excel("matches.xlsx", index=False)

if __name__ == '__main__':

    setup_logger()
    logging.info("Initializing...")

    responses_df = retrieve_data()
    logging.info("Data retrieved")

    inputdata = preprocess_data(responses_df)

    logging.info(f"Preprocess_data: {inputdata}")

    response = match_with_gpt(inputdata)

    logging.info(f"Matching response: {response}")

    postprocess_data(response)

    logging.info("Finished")
