#SlackBot Python Messenger test

import os
from slack_sdk import WebClient

client = WebClient(token=os.getenv('slack_auth'))

#Send a test message
client.chat_postMessage(
    channel="paycom-automation",
    text="Test message from Python script",
    username="Bot User"
)