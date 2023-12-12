import os
import openai
from dotenv import load_dotenv

load_dotenv()
openai.api_key = "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"

def chat_development(user_message):
    conversation = build_conversation(user_message)
    try:
        assistant_message = generate_assistant_message(conversation)
    except openai.error.RateLimitError as e:
        assistant_message = "Rate limit exceeded. Sleeping for a bit..."

    return assistant_message


def build_conversation(user_message):
    return [
        {"role": "system",
         "content": "You are an assistant that generates PowerPoint presentations. I want you to create content for the presentation. Please include relevant information along with elaborations and suggest suitable images or videos. Give atleast 10-15 lines of detailed information for each point."
                    "And the format of the answer must be Slide X(the number of the slide): {title of the content} /n Content: /n content with detailed bullet points."
                    "Keyword: /n Give the most important keyword(within two words) that represents the slide for each one"},
        {"role": "user", "content": user_message}
    ]

def generate_assistant_message(conversation):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=conversation
    )
    return response['choices'][0]['message']['content']