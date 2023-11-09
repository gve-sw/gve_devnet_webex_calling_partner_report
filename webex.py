#!/usr/bin/env python3
"""
Copyright (c) 2023 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

__author__ = "Mark Orszycki <morszyck@cisco.com>, Trevor Maco <tmaco@cisco.com>"
__copyright__ = "Copyright (c) 2023 Cisco and/or its affiliates."
__license__ = "Cisco Sample Code License, Version 1.1"

import base64
import requests
import os
from datetime import datetime
from rich.console import Console
from dotenv import load_dotenv

# Load env variables
load_dotenv()
CCW_CLIENT_ID = os.getenv("CCW_CLIENT_ID")
CCW_CLIENT_SECRET = os.getenv("CCW_CLIENT_SECRET")

base_url = 'https://webexapis.com/v1/'


class WebexCallingInfo:
    def __init__(self, token, id, org_name, console, error_logger):
        self.headers = {'Authorization': f'Bearer {token}'}
        self.console = console if console else Console()
        self.error_logger = error_logger if error_logger else None
        self.error_flag = False
        self.id = id
        self.org_id = None
        self.displayName = org_name
        self.trunks = []
        self.professional_licenses = {}
        self.workspace_licenses = {}
        self.sub_ids = []
        self.sub_start_dates = []
        self.sub_end_dates = []
        self.phone_numbers = []
        self.outgoing_permissions = {}
        self.intercept_settings = {}

    def get_wrapper(self, url, params):
        """
        General function for HTTP GET requests with authentication headers
        """
        # Build Get Request Components
        target_url = f'{base_url}{url}'

        response = requests.get(url=target_url, headers=self.headers, params=params)

        if response.ok:
            return response.json()
        else:
            # Elevated permissions may have expired, attempt to regain permissions with org/id call (once)
            if response.status_code == 403:
                org_details_url = f"{base_url}/organizations/{self.id}"
                response = requests.get(url=org_details_url, headers=self.headers, params={})

                # Make original request
                response = requests.get(url=target_url, headers=self.headers, params=params)

                # If response is ok, permissions reset, return response
                if response.ok:
                    return response.json()

            # Print failure message (error is either not a 403 error or not resolved via 403 fix)
            self.console.print("\n[red]Request FAILED: [/]" + str(response.status_code))
            self.console.print(response.text)
            self.console.print(f"\nAPI Response Headers: {response.headers}")

            # Write Errors to File
            if self.error_logger:
                self.error_logger.error("\nRequest FAILED: " + str(response.status_code))
                self.error_logger.error(response.text)
                self.error_logger.error(f"API Response Headers: {response.headers}")

            self.error_flag = True
            return None

    def get_org_details(self):
        """
        Get Webex Control Hub Org API id and name, necessary for proper scope and permissions to avoid 403 errors
        """
        # Get org name and details
        org_url = f"organizations/{self.id}"
        response = self.get_wrapper(org_url, {})

        if response:
            self.displayName = response['displayName']

    def get_org_id(self):
        """
        Get Webex Control Hub Org id (different from API Org ID). ID is base64 encoded, with additional encoded
        information.
        """
        # Decode the id from base64 (add 2 = signs to ensure divisibility by 4, non-validate param cuts off excess =)
        decoded_id = base64.b64decode(self.id + '==').decode('utf-8')

        # Split the string and retrieve the organization number
        self.org_id = decoded_id.split('/')[-1]

    def get_trunks(self, progress=None):
        """
        Get a list of configured Trunks, retrieve Route Groups associated with Trunks as well
        """
        # Get a List of configured Trunks
        trunk_url = "telephony/config/premisePstn/trunks"
        trunk_params = {'orgId': self.id}

        response = self.get_wrapper(trunk_url, trunk_params)

        if response:
            trunks = response['trunks']

            if progress:
                # Optional progress display
                task = progress.add_task("Find Trunks (and Route Groups)", total=len(trunks), transient=True)

            # Add Trunks to list, determine if trunk is attached to Route Group
            for trunk in trunks:
                trunk_info = {
                    "name": trunk['name'],
                    "id": trunk['id'],
                    'rg_names': []
                }

                # Get Route Group (RG) attached to Trunk
                rg_url = f"telephony/config/premisePstn/trunks/{trunk_info['id']}/usageRouteGroup"
                rg_params = {'orgId': self.id}

                response = self.get_wrapper(rg_url, rg_params)

                if response:
                    RGs = response['routeGroups']

                    # If trunk not attached to RG, skip
                    if len(RGs) == 0:
                        continue

                    trunk_info['rg_names'] = [rg['name'] for rg in RGs]

                    self.trunks.append(trunk_info)

                if progress:
                    progress.update(task, advance=1)

    def get_license_counts(self):
        """
        Get the license count for Webex Calling Professional and Workspace licenses, extract subscription ID's as well
        """
        # Get a list of licenses used in the organization
        license_url = "licenses"
        license_params = {'orgId': self.id}

        response = self.get_wrapper(license_url, license_params)

        if response:
            licenses = response['items']

            # Extract Webex Calling professional and workspaces license specifically
            for license in licenses:
                # Webex Calling - Workspaces case
                if license['name'] == "Webex Calling - Workspaces":
                    # If license already present, sum
                    if 'provisioned' in self.workspace_licenses and 'booked' in self.workspace_licenses:
                        self.workspace_licenses['provisioned'] += license['consumedUnits']
                        self.workspace_licenses['booked'] += license['totalUnits']
                    else:
                        self.workspace_licenses['provisioned'] = license['consumedUnits']
                        self.workspace_licenses['booked'] = license['totalUnits']

                    # Append Subscription ID
                    if 'subscriptionId' in license and license['subscriptionId'] != '':
                        # Avoid adding duplicates if subscription ID already present
                        if license['subscriptionId'] not in self.sub_ids:
                            self.sub_ids.append(license['subscriptionId'])

                # Webex Calling - Professional case
                if license['name'] == "Webex Calling - Professional":
                    # If license already present, sum
                    if 'provisioned' in self.professional_licenses and 'booked' in self.professional_licenses:
                        self.professional_licenses['provisioned'] += license['consumedUnits']
                        self.professional_licenses['booked'] += license['totalUnits']
                    else:
                        self.professional_licenses['provisioned'] = license['consumedUnits']
                        self.professional_licenses['booked'] = license['totalUnits']

                    # Append Subscription ID
                    if 'subscriptionId' in license and license['subscriptionId'] != '':
                        # Avoid adding duplicates if subscription ID already present
                        if license['subscriptionId'] not in self.sub_ids:
                            self.sub_ids.append(license['subscriptionId'])

    def get_license_dates(self):
        """
        Get License Data from CCW GetLicenseData API. Returns Subscription Start and End Date
        """
        # Get CCW Token
        url = "https://id.cisco.com/oauth2/default/v1/token"
        headers = {
            'accept': "application/json",
            'content-type': "application/x-www-form-urlencoded",
            'cache-control': "no-cache"
        }
        payload = "client_id=" + CCW_CLIENT_ID + \
                  "&client_secret=" + CCW_CLIENT_SECRET + \
                  "&grant_type=client_credentials"

        response = requests.request("POST", url, data=payload, headers=headers)

        if response.status_code == 200:
            # Successful token response, get subscription dates
            access_token = response.json()['access_token']

            url = "https://apix.cisco.com/commerce/ORDER/v2/sync/getSubscriptionDetails"
            headers = {
                'authorization': "Bearer " + access_token,
                'accept': "application/json",
                'content-type': "application/json",
                'cache-control': "no-cache"
            }

            for sub_id in self.sub_ids:
                params = {'subscriptionId': sub_id}
                response = requests.get(url, params=params, headers=headers)

                if response.status_code == 200:
                    subscription_info = response.json()['ShowPurchaseOrder']['value']['dataArea']

                    # Check if subscription found (error code GSA003) or some other error not present
                    if 'SUCCESS' in \
                            subscription_info['show']['value']['responseCriteria'][0]['value']['responseExpression'][
                                'value']['value']:
                        # Extract Start and End Dates
                        start_date = subscription_info['purchaseOrder'][0]['value']['purchaseOrderHeader'][
                            'value']['extension'][0]['ciscoExtensionArea']['subscriptionDetail']['value'][
                            'durationAndTerm']['value']['startDateTime']['value']
                        end_date = subscription_info['purchaseOrder'][0]['value']['purchaseOrderHeader'][
                            'value']['extension'][0]['ciscoExtensionArea']['subscriptionDetail']['value'][
                            'durationAndTerm']['value']['endDateTime']['value']

                        # Change formatting of dates to be more readable
                        date_obj = datetime.strptime(start_date, '%Y-%m-%dT%H:%M:%S.%f%z')
                        formatted_start_date = date_obj.strftime('%m/%d/%Y')

                        self.sub_start_dates.append(formatted_start_date)

                        date_obj = datetime.strptime(end_date, '%Y-%m-%dT%H:%M:%S.%f%z')
                        formatted_end_date = date_obj.strftime('%m/%d/%Y')

                        self.sub_end_dates.append(formatted_end_date)
                    else:
                        self.sub_start_dates.append('Unknown')
                        self.sub_end_dates.append('Unknown')

                else:
                    self.console.print("[red]Request FAILED: " + str(response.status_code))
                    self.console.print(response.text)

                    # Write Errors to File
                    if self.error_logger:
                        self.error_logger.error("\nRequest FAILED: " + str(response.status_code))
                        self.error_logger.error(response.text)

                    self.error_flag = True
        else:
            self.console.print("[red]Request FAILED: " + str(response.status_code))
            self.console.print(response.text)

            # Write Errors to File
            if self.error_logger:
                self.error_logger.error("\nRequest FAILED: " + str(response.status_code))
                self.error_logger.error(response.text)

            self.error_flag = True

    def get_phone_numbers(self):
        """
        Determine list of phone numbers assigned to customer org, extract information about the number, it's owner,
        location, etc.
        """
        # Get a list of phone numbers used in the organization
        numbers_url = "telephony/config/numbers"
        numbers_params = {'orgId': self.id}

        response = self.get_wrapper(numbers_url, numbers_params)

        if response:
            numbers = response['phoneNumbers']

            # Iterate through numbers across an organization, extract out relevant fields
            for number in numbers:
                number_info = {
                    'phone_number': '',
                    'main_number': '',
                    'extension': '',
                    'location': '',
                    'owner': '',
                    'owner_id': '',
                    'status': ''
                }

                if 'phoneNumber' in number:
                    number_info['phone_number'] = number['phoneNumber'].replace('+', '') if number['phoneNumber'] else \
                        ''

                if 'mainNumber' in number:
                    number_info['main_number'] = 'Main' if number['mainNumber'] else ''

                if 'extension' in number:
                    number_info['extension'] = number['extension']

                if 'location' in number:
                    number_info['location'] = number['location']['name']

                if 'owner' in number and number['owner']['type'] == 'PEOPLE':
                    # Only add owner id's for people (exclude Voice mail groups, Call Attendants, etc.)
                    number_info['owner'] = number['owner']['firstName'] + ' ' + number['owner']['lastName']

                    # Store owner id for additional permissions settings attached to person and number
                    if number['owner']['id']:
                        number_info['owner_id'] = number['owner']['id']

                if 'state' in number:
                    number_info['status'] = 'Active' if number['state'] == 'ACTIVE' else 'Not Applicable'

                # Append to list
                self.phone_numbers.append(number_info)

    def get_outbound_permissions(self, progress=None):
        """
        Get outbound calling permissions for each user associated to a phone number
        """
        if progress:
            # Optional progress display
            task = progress.add_task("Find Outbound Calling Permissions",
                                                              total=len(self.phone_numbers), transient=True)

        # Get the outbound call permissions for each user
        for phone_number in self.phone_numbers:
            # If the phone number is assigned to a user, grab the permissions by user id
            if phone_number['owner_id'] != '':
                permissions_url = f"people/{phone_number['owner_id']}/features/outgoingPermission"
                permissions_params = {'orgId': self.id}

                response = self.get_wrapper(permissions_url, permissions_params)

                if response:
                    number_permissions = {
                        'internal': '',
                        'toll_free': '',
                        'national': '',
                        'international': '',
                        'operator_assistance': '',
                        'chargeable_directory_assistance': '',
                        'special_services_1': '',
                        'special_services_2': '',
                        'premium_services_1': '',
                        'premium_services_2': ''
                    }

                    # Custom settings enabled, extract individual settings
                    if response['useCustomEnabled']:
                        number_permissions['outgoing_call_permissions'] = 'Custom Settings'

                        permissions = response['callingPermissions']
                        for permission in permissions:
                            if permission['callType'] == 'INTERNAL_CALL':
                                number_permissions['internal'] = permission['action'].title()

                            if permission['callType'] == 'TOLL_FREE':
                                number_permissions['toll_free'] = permission['action'].title()

                            if permission['callType'] == 'NATIONAL':
                                number_permissions['national'] = permission['action'].title()

                            if permission['callType'] == 'INTERNATIONAL':
                                number_permissions['international'] = permission['action'].title()

                            if permission['callType'] == 'OPERATOR_ASSISTED':
                                number_permissions['operator_assistance'] = permission['action'].title()

                            if permission['callType'] == 'CHARGEABLE_DIRECTORY_ASSISTED':
                                number_permissions['chargeable_directory_assistance'] = permission['action'].title()

                            if permission['callType'] == 'SPECIAL_SERVICES_I':
                                number_permissions['special_services_1'] = permission['action'].title()

                            if permission['callType'] == 'SPECIAL_SERVICES_II':
                                number_permissions['special_services_2'] = permission['action'].title()

                            if permission['callType'] == 'PREMIUM_SERVICES_I':
                                number_permissions['premium_services_1'] = permission['action'].title()

                            if permission['callType'] == 'PREMIUM_SERVICES_II':
                                number_permissions['premium_services_2'] = permission['action'].title()
                    else:
                        # Default settings case, no need to record settings
                        number_permissions['outgoing_call_permissions'] = 'Default Settings'

                    # Store settings in dictionary, with key of owner id for easy lookup
                    self.outgoing_permissions[phone_number['owner_id']] = number_permissions

            if progress:
                progress.update(task, advance=1)

    def get_intercept_settings(self, progress=None):
        """
        Get outbound call intercept settings for each user associated to a phone number
        """
        if progress:
            # Optional progress display
            task = progress.add_task("Find Call Intercept Settings",
                                                              total=len(self.phone_numbers), transient=True)
        # Get Outbound Intercept settings for users
        for phone_number in self.phone_numbers:
            # If the phone number is assigned to a user, grab the permissions by user id
            if phone_number['owner_id'] != '':
                intercept_url = f"people/{phone_number['owner_id']}/features/intercept"
                intercept_params = {'orgId': self.id}

                response = self.get_wrapper(intercept_url, intercept_params)

                if response:
                    current_intercept_settings = {}

                    # Extract Call Intercept settings
                    if response['enabled']:
                        current_intercept_settings['call_intercept'] = 'Enable'

                        outgoing_type = response['outgoing']['type']
                        if outgoing_type == 'INTERCEPT_ALL':
                            current_intercept_settings['outgoing_permissions'] = 'Intercept All Outgoing Calls'
                        elif outgoing_type == 'ALLOW_LOCAL_ONLY':
                            current_intercept_settings['outgoing_permissions'] = 'Allow Only National Outgoing Calls'
                    else:
                        current_intercept_settings['call_intercept'] = 'Disable'
                        current_intercept_settings['outgoing_permissions'] = ''

                    # Store settings in dictionary, with key of owner id for easy lookup
                    self.intercept_settings[phone_number['owner_id']] = current_intercept_settings

            if progress:
                progress.update(task, advance=1)
