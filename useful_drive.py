import os.path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# If modifying these SCOPES_DRIVE, delete the file token.json.
# More information in https://developers.google.com/drive/api/v3/about-auth
SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive']


def initialize_drive():
    """
    ->Function to get/create token.

    return: creds: (String) Is a token.
    """

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token_drive.json'):
        creds = Credentials.from_authorized_user_file('token_drive.json', SCOPES_DRIVE)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES_DRIVE)
            creds = flow.run_local_server(port=2)
        # Save the credentials for the next run
        with open('token_drive.json', 'w') as token:
            token.write(creds.to_json())

    return creds

def create_permission(service, fileId, email, role = 'writer', type = 'user', message = '' ,transferOwnership = False):
    """
    -> Function to create permission file in Google Drive.
    ATTENTION: More information in https://developers.google.com/drive/api/v3/reference/permissions/create

    :param service: (String) Is a token. 
    :param fileId: (String) The ID of the file ir shared drive.
    :param role: (String) The role granted by this permission. Valid values are: owner, organizer, fileOrganizer, write, commenter or reader.
    :param type: (String) The type of the grantee. Valid values are: user, group, domain or anyone.
    :param message: (String)(Option) Is a plain text custom message to incluide in the notification email.
    :param transferOwnership: (Boolean) Transfers the ownership if you set True.
    """

    send_email= {
                'role': role,
                'type': type,
                'emailAddress': email
            }

    if len(message) > 1: 
        response_permission = service.permissions().create(
            fileId = fileId, emailMessage=message, transferOwnership = transferOwnership, body=send_email).execute()
    else:
        response_permission = service.permissions().create(
        fileId = fileId, transferOwnership = transferOwnership, body=send_email).execute()