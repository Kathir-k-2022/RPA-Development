"""
Author - Srikanth Koorma
Partner Solution Desk, Automation Anywhere

Packages to be installed: pandas, xlsxwriter & requests

Use below lines in a command prompt to install these packages
pip install pandas
pip install xlsxwriter
pip install requests
"""

# Packages to be imported
import io
import json
import requests
import pandas as pd
import zipfile as zf
from datetime import datetime
from pandas import ExcelWriter

# Uncomment below line to execute this script in cmd
# import sys

# Required variables
line_num = 0
detailed_dict = {}
cmd_dict = {}
consider_disabled = True


# Read data from file
def read_data_from_file(file_path):
    # Open file with utf-8 encoding
    with open(file_path, encoding='utf8') as f:
        file_data = f.read()
        f.close()
    # Returning file contents
    return file_data


# Get token status if its valid or invalid
def token_status(cr_url, user_token):
    headers = {"accept": "application/json"}
    response = requests.get(cr_url + '/v1/authentication/token?token=' + user_token, headers=headers)

    # Checking if response status is 200
    if response.status_code == 200:
        json_obj = response.json()
        return str(json_obj.get('valid'))

    else:
        # Returning error json object
        error_json_obj = response.json()
        return error_json_obj


def generate_token(cr_url, username, password, api_key):
    if password == '' and api_key != '':
        # Api key is not null and password is null
        data = '{ \"username\": \"' + str(username) + '\", \"apiKey\": \"' + str(api_key) + '\"}'

    elif password != '' and api_key == '':
        # Password is not null and api key is null
        data = '{ \"username\": \"' + str(username) + '\", \"password\": \"' + str(password) + '\"}'

    elif password != '' and api_key != '':
        # Password and api key both are not null
        data = '{ \"username\": \"' + str(username) + '\", \"password\": \"' + str(password) + '\"}'

    else:
        # Password and api key both are null
        return json.dumps({'code': 'user.credentials.empty', 'details': '',
                           'message': 'Both password and api key are null. Please provide either one of them to generate user token'})

    headers = {'Content-type': 'application/json', 'Accept': 'application/json'}
    response = requests.post(cr_url + '/v1/authentication', data=data, headers=headers)

    # Checking if response status is 200
    if response.status_code == 200:
        json_obj = response.json()

        # Returning user token
        return json_obj.get('token')
    else:
        error_json_obj = response.json()

        # Returning error json object
        return error_json_obj


# Get the parent node for children
def get_parent(node, cond):
    # Checking if command is a loop start command
    if cond == 'Loop' and str(node.get('Command')).lower() == 'loop.commands.start':
        # Returning parent line number
        return node.get('Line')
    else:
        # Get parent
        return get_parent(node.get('Parent'), cond)


# Get children count for the parent node
def get_children_count(node, count):
    # Looping each child from children
    for sub_node in node.get('Children'):
        # Checking if node child is a step command as this would be counted as children of node
        if sub_node.get('Command').lower() in ['step']:
            count = get_children_count(sub_node, count + 1)

        # Checking if command is disabled and should be considered
        elif not (sub_node.get('Disabled') and consider_disabled):
            count = count + 1

    # Returning children count
    return count


# Update cmd data
def update_cmd_dict(node):
    # Checking id node is a dictionary object
    if type(node) == dict:
        if node.get('disabled'):
            # Assigning command disabled
            command_package_disabled = node.get('packageName') + '|' + node.get('commandName') + '|disabled'
        else:
            # Assigning command enabled
            command_package_disabled = node.get('packageName') + '|' + node.get('commandName') + '|enabled'

        if not cmd_dict.get(command_package_disabled):
            # Incrementing command enabled count
            cmd_dict[command_package_disabled] = 1
        else:
            # Incrementing command disabled count
            cmd_dict[command_package_disabled] = cmd_dict.get(command_package_disabled) + 1

        if node.get('children'):
            # Looping each sub node in children
            for json_sub_node in node.get('children'):
                # Running this to update cmd data
                update_cmd_dict(json_sub_node)

        if node.get('branches'):
            # Looping each sub node in branches
            for json_sub_node in node.get('branches'):
                # Running this to update cmd data
                update_cmd_dict(json_sub_node)

    elif type(node) == list:
        # Looping each node from list
        for json_sub_node in node:
            # Running this to update cmd data
            update_cmd_dict(json_sub_node)


# Update nodes from the parent node
def update_nodes(node):
    global line_num

    # Incrementing line number
    line_num = line_num + 1

    # Inserting data into node with uid value as a key
    detailed_dict[node.get('uid')] = {'Line': line_num, 'Package': node.get('packageName'),
                                      'Command': node.get('commandName'), 'Disabled': node.get('disabled'),
                                      'Children': [], 'Branches': [], 'Parent': {}, 'Dependent Line': '',
                                      'Block Lines': 0}

    # Checking if node has any children
    if node.get('children'):
        # Looping each sub command from children
        for json_sub_node in node.get('children'):
            # Running this to update node data
            update_nodes(json_sub_node)

    # Checking if node has any branches
    if node.get('branches'):
        # Looping each sub command from branches
        for json_sub_node in node.get('branches'):
            # Running this to update node data
            update_nodes(json_sub_node)


# Updating children and branches for parent nodes
def update_children_and_branches(node):
    # Checking if node has any children
    if node.get('children'):
        # Looping each sub command from children
        for json_sub_node in node.get('children'):
            # Saving the parent line number
            detailed_dict.get(json_sub_node.get('uid'))['Parent'] = detailed_dict.get(node.get('uid'))

            # Appending children elements to parent command
            detailed_dict.get(node.get('uid')).get('Children').append(detailed_dict.get(json_sub_node.get('uid')))

            # Running this to update children and branches for all the nodes
            update_children_and_branches(json_sub_node)

    # Checking if node has any branches
    if node.get('branches'):
        # Looping each sub command from branches
        for json_sub_node in node.get('branches'):
            # Saving the parent line number
            detailed_dict.get(json_sub_node.get('uid'))['Parent'] = detailed_dict.get(node.get('uid'))

            # Appending branch elements to parent command
            detailed_dict.get(node.get('uid')).get('Branches').append(detailed_dict.get(json_sub_node.get('uid')))

            # Running this to update children and branches for all the nodes
            update_children_and_branches(json_sub_node)


# Export bot from control room
def export_bot(cr_url, user_token, export_file_name, bot_id, executor_username):
    if export_file_name == '':
        data = '{\"name\": \"' + 'Export.' + datetime.today().strftime('%Y%m%d_%H%M%S') + '.' + str(executor_username) + '\", \"fileIds\": [' + str(
            bot_id) + '], \"includePackages\": false}'
    else:
        data = '{\"name\": \"' + str(export_file_name) + '\", \"fileIds\": [' + str(
            bot_id) + '], \"includePackages\": false}'

    headers = {"X-Authorization": user_token, 'Content-type': 'application/json', 'Accept': 'text/plain'}
    response = requests.post(cr_url + '/v2/blm/export', data=data, headers=headers)

    # Checking if response status is 202
    if response.status_code == 202:
        json_obj = response.json()
        return json_obj.get('requestId')

    else:
        error_json_obj = response.json()
        # Returning error json object
        return error_json_obj


# Get bot export status
def bot_export_status(cr_url, request_id, user_token):
    headers = {"X-Authorization": user_token, "accept": "application/json"}
    response = requests.get(cr_url + '/v2/blm/status/' + request_id, headers=headers)

    # Checking if response status is 200
    if response.status_code == 200:
        json_obj = response.json()

        # Checking bot export status
        if json_obj.get('status').lower() == 'completed':
            # Returning download file id attribute
            return json_obj.get('downloadFileId')

        else:
            # Returning wait
            return 'wait'
    else:
        error_json_obj = response.json()

        # Returning error json object
        return error_json_obj


# Download exported bot files and export them into local folder
def download_file(cr_url, download_id, user_token, folder_path):
    headers = {"X-Authorization": user_token,
               "accept": "*/*", "accept-encoding": "gzip;deflate;br"}
    response = requests.get(cr_url + '/v2/blm/download/' + download_id, headers=headers)

    # Checking if response status is 200
    if response.status_code == 200:
        z = zf.ZipFile(io.BytesIO(response.content))
        z.extractall(folder_path)

        # Extracting zip file contents
        return "ok"

    else:
        error_json_obj = response.json()

        # Returning error json object
        return error_json_obj


# Generate analysis file
def generate_analysis(file_path):
    # Loading the parent node which is read from file
    json_dict = json.loads(read_data_from_file(file_path)).get('nodes')

    # Running this to update cmd dictionary
    update_cmd_dict(json_dict)

    temp_dict = {}

    # Looping each element from dictionary
    for i, item in enumerate(cmd_dict.items()):
        package, command, status = item[0].split("|")

        # Saving command analysis to a dictionary
        temp_dict[i] = {"Package": package, "Command": command, "Status": status, "Frequency": item[1]}

    # Loading command analysis data from dictionary to pandas dataframe
    df_cmd = pd.DataFrame.from_dict(temp_dict, orient="index")
    df_cmd = df_cmd.sort_values(['Package', 'Command', 'Status', 'Frequency'], ascending=[True, True, True, True])

    # Looping each sub node from json dictionary
    for node_json_dict in json_dict:
        # Updating nodes
        update_nodes(node_json_dict)

        # Updating children and branch nodes
        update_children_and_branches(node_json_dict)

    # Looping each command from dictionary object
    for uid in detailed_dict:

        # Get parents for loop continue and break
        if str(detailed_dict.get(uid).get('Command')).lower() in ['loop.commands.break', 'loop.commands.continue']:
            # Get parent and save it to the dependent line
            detailed_dict.get(uid)['Dependent Line'] = get_parent(detailed_dict.get(uid).get('Parent'), 'Loop')

        # Check if current command is not branches of parent as these branch commands are already calculated in parent
        if not str(detailed_dict.get(uid).get('Command')).lower() in ['else', 'elseif', 'catch', 'finally',
                                                                      'step'] and not (
                detailed_dict.get(uid).get('Disabled') and consider_disabled):

            # Check if current node has any children
            if len(detailed_dict.get(uid).get('Children')) > 0:

                # Loop each child from children
                for sub_node in detailed_dict.get(uid).get('Children'):

                    # Checking if node child is a step command as this would be counted as children of node
                    if sub_node.get('Command').lower() in ['step'] and not (
                            sub_node.get('Disabled') and consider_disabled):

                        # Adding children of step command as node children
                        detailed_dict.get(uid)['Block Lines'] = get_children_count(sub_node, detailed_dict.get(uid).get(
                            'Block Lines') + 1)

                    elif not (sub_node.get('Disabled') and consider_disabled):
                        # Adding block line
                        detailed_dict.get(uid)['Block Lines'] = detailed_dict.get(uid).get('Block Lines') + 1

            # Check if current node has any branches
            if len(detailed_dict.get(uid).get('Branches')) > 0:

                # Looping each child from branches
                for sub_node in detailed_dict.get(uid).get('Branches'):

                    # Checking sub command is disabled or not
                    if not (sub_node.get('Disabled') and consider_disabled):

                        # Adding block line
                        detailed_dict.get(uid)['Block Lines'] = detailed_dict.get(uid).get('Block Lines') + 1

                        # Looping each child from children
                        for sub_sub_node in sub_node.get('Children'):

                            # Checking if sub command is a step and not disabled
                            if sub_sub_node.get('Command').lower() in ['step'] and not (
                                    sub_sub_node.get('Disabled') and consider_disabled):

                                # Get children count and add line count to block lines
                                detailed_dict.get(uid)['Block Lines'] = detailed_dict.get(uid).get('Block Lines') + 1
                                detailed_dict.get(uid)['Block Lines'] = get_children_count(sub_sub_node, detailed_dict.get(uid).get('Block Lines'))
                            else:
                                # Adding block line
                                detailed_dict.get(uid)['Block Lines'] = detailed_dict.get(uid).get('Block Lines') + 1

    temp_dict = {}

    # Looping each command from dictionary
    for i, item in enumerate(detailed_dict.items()):
        # Saving command data to dictionary
        if item[1].get('Parent'):
            temp_dict[i] = {"Line": item[1].get('Line'), "Package": item[1].get('Package'),
                            "Command": item[1].get('Command'),
                            "Disabled": item[1].get('Disabled'), "Parent": item[1].get('Parent').get('Line'),
                            "Dependent Line": item[1].get('Dependent Line'), "Block Lines": item[1].get('Block Lines')}
        else:
            temp_dict[i] = {"Line": item[1].get('Line'), "Package": item[1].get('Package'),
                            "Command": item[1].get('Command'),
                            "Disabled": item[1].get('Disabled'), "Parent": 0,
                            "Dependent Line": item[1].get('Dependent Line'), "Block Lines": item[1].get('Block Lines')}

    # Get dictionary data and store in pandas dataframe
    df_detailed = pd.DataFrame.from_dict(temp_dict, orient="index")

    # Write excel file to local folder
    with ExcelWriter(file_path + '.xlsx', engine='xlsxwriter') as writer:
        # Writing command analysis data to excel sheet
        df_cmd.to_excel(writer, encoding='utf-8', index=False, sheet_name="Command Analysis")

        # Writing detailed analysis data to excel sheet
        df_detailed.to_excel(writer, encoding='utf-8', index=False, sheet_name="Detailed Analysis")


# Perform code review
def code_review(args):
    try:
        cr_url = args.get('cr_url')
        user_token = args.get('user_token')
        bot_id = args.get('bot_id')
        export_file_name = args.get('export_file_name')
        folder_path = args.get('folder_path')
        executor_username = args.get('executor_username')

        if str(cr_url).endswith("/"):
            cr_url = cr_url[:-1]

        # Get token status
        tok_status = token_status(cr_url, user_token)

        # Checking token status
        if type(tok_status) == str and tok_status == 'false':
            return json.dumps({'code': 'user.token.invalid', 'details': '',
                               'message': 'Given user token is invalid. Please generate a new one.'})
        elif type(tok_status) == dict:
            return json.dumps(tok_status)

        # Export bot from control room
        request_id = export_bot(cr_url, user_token, export_file_name, bot_id, executor_username)

        # Checking if bot export has started
        if type(request_id) == str and request_id != '':
            download_id = 'wait'

            # Waiting for bot export to be completed
            while type(download_id) == str and download_id == 'wait':
                # Get bot export status
                download_id = bot_export_status(cr_url, request_id, user_token)

            # Checking if bot export is completed
            if type(download_id) == str and download_id != 'wait' and download_id != '':
                # Downloading file
                download_obj = download_file(cr_url, download_id, user_token, folder_path)

                # Checking if download is success
                if type(download_obj) == str and download_obj == 'ok':
                    json_obj = json.loads(read_data_from_file(folder_path + "\\manifest.json"))

                    # Looping all bot files
                    for file in json_obj.get('files'):
                        # Checking if file is a bot file
                        if file.get('contentType') in ['application/vnd.aa.taskbot']:
                            # Generating analysis file
                            generate_analysis(folder_path + '\\' + file.get('path'))

                    # Returning success
                    return json.dumps({'code': 'ok', 'details': '',
                                       'message': 'Generated bot analysis report is available in respective bot folders at ' + folder_path})
                else:
                    # Returning error
                    return json.dumps(download_obj)
            else:
                # Returning error
                return json.dumps(download_id)
        else:
            # Returning error
            return json.dumps(request_id)
    except Exception as err:
        return json.dumps({'code': 'python.exception', 'details': '', 'message': str(err)})


# Uncomment below lines to execute in cmd
"""
args = sys.argv[1:]

cr_url = args[0]
username = args[1]
password = args[2]
api_key = args[3]
bot_id = args[4]
export_file_name = args[5]
folder_path = args[6]
executor_username = args[7]

user_token = generate_token(cr_url, username, password, api_key)

if type(user_token) == str:
    print(code_review({'cr_url': cr_url, 'user_token': user_token, 'bot_id': bot_id, 'export_file_name': export_file_name, 'folder_path': folder_path, 'executor_username': executor_username}))
elif type(user_token) == dict:
    print(json.dumps(user_token))
"""
