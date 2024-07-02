import sys
from openai import OpenAI
import time
import re

def EngSummary(long_sentence):
    # print("long: ", long_sentence)
    client = OpenAI(api_key="your OPENAI API Key Here!")

    assistant_ID="your assistant ID Here!"
    # assistant = client.beta.assistants.create(
    #     name="English Zipper",
    #     description="You are an IT information and communication expert. You can dramatically compress English sentences in a sophisticated way without hurting the meaning of them. If necessary, long words can be replaced by short abbreviations that are widely known",
    #     model="gpt-4",
    #     tools=[{"type": "code_interpreter"}]
    # )

    thread_ID="your thread ID Here!"
    # thread = client.beta.threads.create()

    # Step 3: Add a Message to a Thread
    message = client.beta.threads.messages.create(
        # thread_id=thread.id,
        thread_id=thread_ID,
        role="user",
        content=long_sentence
    )

    #Step 4: Run the Assistant
    run = client.beta.threads.runs.create(
        # thread_id=thread.id,
        thread_id=thread_ID,
        assistant_id=assistant_ID
    )

    # Check the Run Status

    while True:
        run = client.beta.threads.runs.retrieve(
            # thread_id=thread.id,
            thread_id=thread_ID,
            run_id=run.id
            )
        if run.status == "completed":
            break
        else:
            time.sleep(2)
        # print(run)

    # Step 6: Display the Assistant's Response
    messages = client.beta.threads.messages.list(
        # thread_id=thread.id
        thread_id=thread_ID
    )

    # print(messages.data)

    short_sentence = messages.data[0].content[0].text.value
    # print("short >> ", short_sentence)
    return short_sentence

if __name__ == "__main__":
    # Extract the input sentence from command-line arguments
    input_sentence = sys.argv[1]
    output_sentence = EngSummary(input_sentence)
    # return output_sentence
    print(output_sentence)

# resultmsg = EngSummary("Cisco Systems (NASDAQ:CSCO) lowers its annual guidance and outlines plans to slash headcount, as executives at the network equipment manufacturer warn of weak future demand")
# print(resultmsg)