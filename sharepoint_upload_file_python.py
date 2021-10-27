#pip install Office365-REST-Python-Client
#pip install git+https://github.com/vgrem/Office365-REST-Python-Client.git

# courtesy: https://stackoverflow.com/questions/59979467/accessing-microsoft-sharepoint-files-and-data-using-python

#Importing required libraries

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 

#Constrtucting SharePoint URL and credentials 

sharepoint_base_url = 'https://mycompany.sharepoint.com/teams/sharepointname/'
sharepoint_user = 'user'
sharepoint_password = 'pwd'
folder_in_sharepoint = '/teams/sharepointname/Shared%20Documents/YourFolderName/'

#Constructing Details For Authenticating SharePoint

auth = AuthenticationContext(sharepoint_base_url)

auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
ctx = ClientContext(sharepoint_base_url, auth)
web = ctx.web
ctx.load(web)
ctx.execute_query()
print('Connected to SharePoint: ',web.properties['Title'])

#Read filename
fileName = 'C:\\path\\file.xlsx'   

with open(fileName, 'rb') as content_file:
    file_content = content_file.read()

name = os.path.basename(fileName)

list_title = "Documents"

target_list = ctx.web.lists.get_by_title(list_title)
info = FileCreationInformation()

libraryRoot = ctx.web.get_folder_by_server_relative_url(folder_url_shrpt)

target_file = libraryRoot.upload_file(name, file_content).execute_query()
print("File has been uploaded to url: {0}".format(target_file.serverRelativeUrl))
