import base64
import os
import smtplib
import msal
import requests
import pandas as pd

def init_ignore_list(dump_dir: str) -> list:
    """
    Initialize the ignore list from a CSV file.
    Args:
        dump_dir (str): The directory where the CSV file is located.
    Returns:
        list: The ignore list loaded from the CSV file.
    """
    ignore_filename = os.path.join(dump_dir, 'ignore_list.csv')
    if not os.path.exists(ignore_filename):
        return []
    df = pd.read_csv(ignore_filename)
    return list(df.iloc[:, 0])


def add_to_list_file(msg:str, filename:str):
    """
    Add a message to a list file.

    Args:
        msg (str): The message to be added.
        filename (str): The path to the file.

    Returns:
        None
    """
    if os.path.exists(filename):
        df = pd.read_csv(filename)
    else:
        df = pd.DataFrame([], columns=['msg_id'])
    df.loc[len(df)] = [msg]
    df.to_csv(filename, index=False)


def add_ignore_list(msg:str, dump_dir:str):
    """
    Add a message to the ignore list file.

    Args:
        msg (str): The message to be added to the ignore list.
        dump_dir (str): The directory where the ignore list file is located.

    Returns:
        None
    """
    ignore_filename = os.path.join(dump_dir, 'ignore_list.csv')
    add_to_list_file(msg, ignore_filename)


def get_auth_token(client_id:str, tenant_id:str, username:str, password:str) -> str:
    """
    Retrieves an authentication token using the provided client ID, tenant ID, username, and password.

    Parameters:
    client_id (str): The client ID of the application.
    tenant_id (str): The tenant ID of the Azure AD.
    username (str): The username of the user.
    password (str): The password of the user.

    Returns:
    str: The access token for authentication.
    """
    authority = f'https://login.microsoftonline.com/{tenant_id}'
    scope = ["https://graph.microsoft.com/.default"]

    app = msal.PublicClientApplication(
        client_id,
        authority=authority,
    )

    result = app.acquire_token_silent(scope, account=None)
    if not result:
        result = app.acquire_token_by_username_password(username=username, password=password, scopes=scope)
    if "access_token" in result:
        access_token = result['access_token']
    return access_token


def get_folder_id(access_token:str, folder_name:str) -> str:
    """
    Retrieves the ID of a mail folder with the specified name.

    Parameters:
    access_token (str): The access token for authentication.
    folder_name (str): The name of the mail folder.

    Returns:
    str or None: The ID of the mail folder if found, None otherwise.
    """
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get('https://graph.microsoft.com/v1.0/me/mailFolders', headers=headers)
    try:
        folders = response.json()['value']
    except KeyError:
        return None
    
    for folder in folders:
        if folder['displayName'] == folder_name:
            return folder['id']
    return None


def get_messages(access_token:str, folder_id:str) -> list:
    """
    Retrieves a list of message IDs from a specified folder using the Microsoft Graph API.

    Parameters:
    access_token (str): The access token for authentication.
    folder_id (str): The ID of the folder from which to retrieve the messages.

    Returns:
    list: A list of message IDs.

    Example:
    access_token = 'your_access_token'
    folder_id = 'your_folder_id'
    messages = get_messages(access_token, folder_id)
    """
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(f'https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?$top=999', headers=headers)
    messages = response.json()
    messages_id = []
    for message in messages['value']:
        messages_id.append(message['id'])
    return messages_id


def get_attachments(access_token:str, message_id:str) -> list:
    """
    Retrieves a list of attachments from a specified message using the Microsoft Graph API.

    Parameters:
    access_token (str): The access token for authentication.
    message_id (str): The ID of the message from which to retrieve the attachments.

    Returns:
    list: A list of attachments.

    Example:
    access_token = 'your_access_token'
    message_id = 'your_message_id'
    attachments = get_attachments(access_token, message_id)
    """
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments", headers=headers)
    attachments = response.json()
    return attachments



if __name__ == '__main__':
    
    client_id = os.getenv('CLIENT_ID')
    tenant_id = os.getenv('TENANT_ID')
    username = os.getenv('USERNAME')
    password = os.getenv('PASSWORD')
    server = os.getenv('SMTP_SERVER')
    port = os.getenv('SMTP_PORT')
    access_token = get_auth_token(client_id, tenant_id, username, password) 
    folder_id = get_folder_id(access_token, 'Вхідні') 
    messages_id = get_messages(access_token, folder_id)
    attachments = get_attachments(access_token, messages_id[0])
    

    
