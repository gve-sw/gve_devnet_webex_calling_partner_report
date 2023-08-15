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

import json
import os
import shutil
import smtplib
import sys
import time
import logging
from datetime import date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests
from requests_oauthlib import OAuth2Session
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress

import config
from webex import WebexCallingInfo

# Refreshing token URL
TOKEN_URL = 'https://api.ciscospark.com/v1/access_token'

REPORT_TEMPLATE = "calling_report_template.xlsx"

# Rich console instance
console = Console()

# Specify list of Org names to specifically process (only these orgs will be processed)
ORGS = []


def custom_logger(date):
    """
    Define custom logger for each report run, writes logs to date stamped file. Creates or returns existing logger. Logs for the current day overwritten with each run!
    :param date: Report date
    :return: logger instance
    """
    logger = logging.getLogger(date)

    if not logger.handlers:
        logFile = os.path.join('./logs/', 'report_errors_' + date + '.log')

        os.makedirs(os.path.dirname(logFile), exist_ok=True)

        # Create logger
        my_handler = logging.FileHandler(filename=logFile, mode='w')
        my_handler.setLevel(logging.INFO)
        logger.addHandler(my_handler)

        logger.setLevel(logging.INFO)

    return logger


def refresh_token(tokens):
    """
    Refresh Webex token if primary token is expired (assumes refresh token is valid)
    :param tokens: Primary and Refresh Tokens
    :return: New set of tokens
    """
    refresh_token = tokens['refresh_token']
    extra = {
        'client_id': config.WEBEX_CLIENT_ID,
        'client_secret': config.WEBEX_CLIENT_SECRET,
        'refresh_token': refresh_token,
    }
    auth_code = OAuth2Session(config.WEBEX_CLIENT_ID, token=tokens)
    new_teams_token = auth_code.refresh_token(TOKEN_URL, **extra)

    # store away the new token
    with open('tokens.json', 'w') as json_file:
        json.dump(new_teams_token, json_file)

    console.print("- [green]A new token has been generated and stored in `tokens.json`[/]")
    return new_teams_token['access_token']


def get_customer_orgs(token):
    """
    List all organizations visible to Webex Control Hub user
    :param token: Webex Tokens
    :return: List of organizations the user belongs to (or manages)
    """
    base_url = 'https://webexapis.com/v1/'
    orgs_url = 'organizations'

    # Build Get Request Components
    target_url = f'{base_url}{orgs_url}'

    response = requests.get(url=target_url, headers={'Authorization': f'Bearer {token}'})

    if response.status_code == 200:
        orgs = response.json()['items']

        console.print(f'Found {len(orgs) - 1 if len(orgs) > 1 else len(orgs)} org(s)!')
        return orgs
    else:
        console.print("[red]Request FAILED: " + str(response.status_code))
        console.print(response.text)
        return None


def populate_df(calling_info, report_number):
    """
    Populate dataframe with information gathered from script, data is written to specific DF depending on report,
    DF is ultimately written to Excel file
    :param calling_info: Webex Calling info gathered from script
    :param report_number: Report number which determines the type of report
    :return: Report DF populated with new information from script per org
    """
    rows = []
    if report_number == 1:
        df_row = {'Customer Name': calling_info.displayName, 'Customer Org ID': calling_info.org_id,
                  'Sub-Ref Id(s)': ', '.join(calling_info.sub_ids),
                  'Recent Subscription Start Date': ', '.join(calling_info.sub_start_dates),
                  'Recent Subscription End Date': ', '.join(calling_info.sub_end_dates), 'Booked (TOTAL)': '',
                  'Booked Professional Licenses': '', 'Booked Workspaces': '',
                  'Provisioned (TOTAL)': '', 'Provisioned Professional Licenses': '', 'Provisioned Workspaces': ''}

        if len(calling_info.professional_licenses) != 0:
            df_row['Booked Professional Licenses'] = calling_info.professional_licenses['booked']
            df_row['Provisioned Professional Licenses'] = calling_info.professional_licenses['provisioned']

        if len(calling_info.workspace_licenses) != 0:
            df_row['Booked Workspaces'] = calling_info.workspace_licenses['booked']
            df_row['Provisioned Workspaces'] = calling_info.workspace_licenses['provisioned']

        # Quick summation for total
        sum_1 = int(df_row['Booked Professional Licenses']) if df_row['Booked Professional Licenses'] != '' else 0
        sum_2 = int(df_row['Booked Workspaces']) if df_row['Booked Workspaces'] != '' else 0
        df_row['Booked (TOTAL)'] = sum_1 + sum_2

        sum_1 = int(df_row['Provisioned Professional Licenses']) if df_row[
                                                                        'Provisioned Professional Licenses'] != '' else 0
        sum_2 = int(df_row['Provisioned Workspaces']) if df_row['Provisioned Workspaces'] != '' else 0
        df_row['Provisioned (TOTAL)'] = sum_1 + sum_2

        rows.append(pd.DataFrame([df_row]))

    elif report_number == 2:
        if len(calling_info.phone_numbers) == 0:
            # No numbers present in org, append blank line
            rows.append(
                pd.DataFrame([{'Customer Name': calling_info.displayName, 'Customer Org ID': calling_info.org_id,
                               'Phone Number': '', 'Main Number': '', 'Extension': '', 'Location': '',
                               'Assigned to': '', 'Status': '',
                               'Outgoing Call Permissions': '', 'Internal': '',
                               'Toll-free': '', 'National': '', 'International': '',
                               'Operator Assistance': '', 'Chargeable Directory Assistance': '',
                               'Special Services I': '', 'Special Services II': '', 'Premium Service I': '',
                               'Premium Services II': '', 'Call Intercept': '',
                               'Outgoing Intercept Permissions': ''}]))
        else:
            # Iterate through list of numbers, append to DF
            for number in calling_info.phone_numbers:
                df_row = {'Customer Name': calling_info.displayName, 'Customer Org ID': calling_info.org_id,
                          'Phone Number': number['phone_number'], 'Main Number': number['main_number'],
                          'Extension': number['extension'], 'Location': number['location'],
                          'Assigned to': number['owner'], 'Status': number['status'],
                          'Outgoing Call Permissions': '', 'Internal': '',
                          'Toll-free': '', 'National': '', 'International': '',
                          'Operator Assistance': '', 'Chargeable Directory Assistance': '',
                          'Special Services I': '', 'Special Services II': '', 'Premium Services I': '',
                          'Premium Services II': '', 'Call Intercept': '',
                          'Outgoing Intercept Permissions': ''}

                if number['owner_id'] != '':
                    # Grab outgoing call permission data
                    permissions = calling_info.outgoing_permissions[number['owner_id']]
                    df_row['Outgoing Call Permissions'] = permissions['outgoing_call_permissions']

                    # If custom settings enabled, record those settings
                    if df_row['Outgoing Call Permissions'] == 'Custom Settings':
                        df_row['Internal'] = permissions['internal']
                        df_row['Toll-free'] = permissions['toll_free']
                        df_row['National'] = permissions['national']
                        df_row['International'] = permissions['international']
                        df_row['Operator Assistance'] = permissions['operator_assistance']
                        df_row['Chargeable Directory Assistance'] = permissions['chargeable_directory_assistance']
                        df_row['Special Services I'] = permissions['special_services_1']
                        df_row['Special Services II'] = permissions['special_services_2']
                        df_row['Premium Services I'] = permissions['premium_services_1']
                        df_row['Premium Services II'] = permissions['premium_services_2']

                    # Grab Outgoing intercept permissions
                    intercept = calling_info.intercept_settings[number['owner_id']]
                    df_row['Call Intercept'] = intercept['call_intercept']
                    df_row['Outgoing Intercept Permissions'] = intercept['outgoing_permissions']

                rows.append(pd.DataFrame([df_row]))

    elif report_number == 3:
        if len(calling_info.trunks) == 0:
            # No trunks present in org
            rows.append(
                pd.DataFrame([{'Customer Name': calling_info.displayName, 'Customer Org ID': calling_info.org_id,
                               'TRUNK': '', 'ROUTE GROUP NAME': ''}]))
        else:
            # Iterate through list of trunks
            for trunk in calling_info.trunks:
                rows.append(
                    pd.DataFrame([{'Customer Name': calling_info.displayName, 'Customer Org ID': calling_info.org_id,
                                   'TRUNK': trunk['name'], 'ROUTE GROUP NAME': ','.join(trunk['rg_names'])}]))

    # Append new row(s)
    df = pd.concat(rows, ignore_index=True, sort=False)
    return df


def send_email_with_attachment(attachment_path):
    """
    Send report via email (Outlook supported)
    :param attachment_path: Excel Report file path
    """
    # Create subject
    current_date = date.today()
    date_string = current_date.strftime("%m-%d-%Y")

    subject = f'Webex Calling Report - {date_string}'

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = config.EMAIL_USERNAME
    msg['To'] = ','.join(config.RECIPIENTS)
    msg['Subject'] = subject

    message = 'This is an automated message containing information about managed Webex Calling Customers. Please see ' \
              'the attached Report for more information. '
    # Add the message body
    msg.attach(MIMEText(message, 'plain'))

    # Attach the Excel report
    attachment = open(attachment_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{attachment_path.split("/")[-1]}"')
    msg.attach(part)

    # Establish the SMTP connection and send the email
    with smtplib.SMTP(config.SMTP_DOMAIN, config.SMTP_PORT) as server:
        server.starttls()
        server.login(config.EMAIL_USERNAME, config.EMAIL_PASSWORD)
        server.send_message(msg)

    console.print(f'Email sent successfully to [yellow]{config.RECIPIENTS}[/]')


def generate_calling_report(token):
    """
    Generate various Webex Calling reports, populate respective dataframes with data from Webex API,
    write dataframe(s) to report file
    :param token: Webex tokens
    """
    # Define customer error logger
    current_date = date.today()
    date_string = current_date.strftime("%m-%d-%Y")
    error_logger = custom_logger(date_string)

    console.print(Panel.fit("Generate Webex Calling Report(s)", title="Step 2"))

    # Get a list of customer orgs
    orgs = get_customer_orgs(token)

    if not orgs:
        console.print('[red]No customer orgs found, exiting...[/]')
        return

    # Selectively process orgs (if relevant - useful for debugging or selective processing)
    if len(ORGS) > 0:
        orgs = [org for org in orgs if org['displayName'] in ORGS]

    # Org count - 1 because partner org itself is ignored (see config.py)
    orgs_count = len(orgs) - 1 if len(orgs) > 1 else len(orgs)

    with Progress() as progress:
        overall_progress = progress.add_task("Overall Progress", total=orgs_count, transient=True)
        counter = 1

        for org in orgs:
            # Skip partner org (only process customer orgs)
            if org['displayName'] == config.PARTNER_ORG_NAME:
                continue

            # Initialize object which will hold org-wide calling info for generating various reports
            calling_info = WebexCallingInfo(token, org['id'], org['displayName'], progress.console, error_logger)

            # Get Webex Org Id
            calling_info.get_org_id()

            progress.console.print(
                "\nProcessing Org: [blue]'{}'[/] ({} of {})".format(org['displayName'], str(counter), orgs_count))
            error_logger.info("\nProcessing Org: {} ({} of {})".format(org['displayName'], str(counter), orgs_count))

            # Retrieve various Webex Calling Info for each report, stored within object for particular Org
            ### REPORT 1 ###
            progress.console.print("\n- Generating [blue]Report 1[/]:")
            error_logger.info("\n- Generating Report 1:")

            # Gather Professional and Workspace License Information
            error_logger.info("-- License Counts:")
            calling_info.get_license_counts()
            if len(calling_info.professional_licenses) == 0 and len(calling_info.workspace_licenses) == 0:
                progress.console.print("-- [red]No Webex Calling License information found.[/]")
            else:
                if len(calling_info.professional_licenses) != 0:
                    progress.console.print(
                        f"-- Found [green]Webex Calling Professional Licenses[/]: {calling_info.professional_licenses} ([yellow]Sub-Ref ID(s): {calling_info.sub_ids}[/])")

                if len(calling_info.workspace_licenses) != 0:
                    progress.console.print(
                        f"-- Found [green]Webex Calling Workspace Licenses[/]: {calling_info.workspace_licenses} ([yellow]Sub-Ref ID(s): {calling_info.sub_ids}[/])")

            # Gather License Subscription Information (Start and End Date)
            if config.CCW_INTEGRATION:
                error_logger.info("-- License Subscription Dates:")
                calling_info.get_license_dates()
                if len(calling_info.sub_start_dates) == 0 and len(calling_info.sub_end_dates) == 0:
                    progress.console.print("-- [red]Unable to obtain data from CCW API. Skipping...[/]")
                else:
                    progress.console.print("-- Found the following [green]License Start and License End dates[/]: "
                                           f"{calling_info.sub_start_dates}, {calling_info.sub_end_dates}")

            # Populate DF with report info
            df_report1 = populate_df(calling_info, 1)

            ### REPORT 2 ###
            progress.console.print("\n- Generating [blue]Report 2[/]:")
            error_logger.info("\n- Generating Report 2:")

            # Gather Users and Phone Numbers
            error_logger.info("-- Phone Numbers:")
            calling_info.get_phone_numbers()
            if len(calling_info.phone_numbers) == 0:
                progress.console.print("-- [red]No Webex Calling Phone Numbers provisioned.[/]")
            else:
                progress.console.print(
                    f"-- Found [green]Webex Phone Numbers[/]: {[number['phone_number'] for number in calling_info.phone_numbers if number['phone_number'] != '']}")

                # Gather User Outbound Calling Permissions
                error_logger.info("--- Outbound Calling Permissions:")
                calling_info.get_outbound_permissions(progress)
                progress.console.print(f"--- Found [green]Outbound Permissions[/] for each number")

                # Gather User Outbound Calling Intercept Settings
                error_logger.info("--- Call Intercept Settings:")
                calling_info.get_intercept_settings(progress)
                progress.console.print(f"--- Found [green]Outbound Intercept Settings[/] for each number")

            # Populate DF with report info
            df_report2 = populate_df(calling_info, 2)

            ### REPORT 3 ###
            progress.console.print("\n- Generating [blue]Report 3[/]:")
            error_logger.info("\n- Generating Report 3:")

            # Gather Trunk Information
            error_logger.info("-- Calling Trunks:")
            calling_info.get_trunks(progress)

            if len(calling_info.trunks) == 0:
                progress.console.print("-- [red]No Webex Calling Trunks found.[/]")
            else:
                progress.console.print(
                    f"-- Found [green]Webex Calling Trunks[/]: {calling_info.trunks}")

            # Populate DF with report info
            df_report3 = populate_df(calling_info, 3)

            counter += 1
            progress.update(overall_progress, advance=1)

            # Cleanup Intermediate progress displays (ignore first task -> overall task)
            task_ids = progress.task_ids
            task_ids.pop(0)
            for task_id in task_ids:
                progress.remove_task(task_id)

    console.print(Panel.fit("Saving File (Sending Email)", title="Step 3"))

    if config.CSV_FORMAT:
        # Create 3 CSV reports, combine them into a folder, and zip the folder

        # Define base naming convention for CSV files, create temp directory
        base_name = f"calling_report_{date_string}"
        os.makedirs(base_name, exist_ok=True)

        # Write out to CSV files within temp directory created
        df_report1.to_csv(os.path.join(base_name, f'{base_name}_1.csv'), index=False)
        df_report2.to_csv(os.path.join(base_name, f'{base_name}_2.csv'), index=False)
        df_report3.to_csv(os.path.join(base_name, f'{base_name}_3.csv'), index=False)

        # Zip up the directory
        shutil.make_archive(base_name, 'zip', base_name)

        # Clean up - remove the original directory
        shutil.rmtree(base_name)

        # Destination file = zip file, receive same moving and email treatment
        destination_file = f"{base_name}.zip"

    else:
        # Output reports in condensed Excel format (based on template file)
        destination_file = f"calling_report_{date_string}.xlsx"

        # Create New Report based on template (date-stamped clone)
        shutil.copy(REPORT_TEMPLATE, destination_file)

        with pd.ExcelWriter(destination_file, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df_report1.to_excel(writer, sheet_name='Report #1', startrow=1, index=False, header=False)

        with pd.ExcelWriter(destination_file, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df_report2.to_excel(writer, sheet_name='Report #2', startrow=1, index=False, header=False)

        with pd.ExcelWriter(destination_file, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            df_report3.to_excel(writer, sheet_name='Report #3', startrow=1, index=False, header=False)

    console.print(f'New report file created: `[blue]{destination_file}[/]`')

    if config.DESTINATION_PATH != '' and os.path.exists(config.DESTINATION_PATH):
        # if destination path defined, check if path exists, then move the file if successful
        new_path = shutil.move(destination_file, os.path.join(config.DESTINATION_PATH, destination_file))
        console.print(f"Saving File to [green]{new_path}[/]")

    else:
        # Create the default reports directory if it doesn't exist
        new_folder_path = os.path.join(os.getcwd(), 'reports')
        os.makedirs(new_folder_path, exist_ok=True)

        # Move the Excel file to the destination
        new_path = shutil.move(destination_file, os.path.join(new_folder_path, destination_file))
        console.print(f"Saving File to [green]{new_path}[/]")

    # Note: If Send Email enabled, it's assumed email configured correctly
    if config.SEND_EMAIL:
        # Send Email with report attached, sends email to recipient list configured
        send_email_with_attachment(new_path)


def main():
    console.print(Panel.fit("Webex Partner Calling Report"))

    # If token file already exists, extract existing tokens
    if os.path.exists('tokens.json'):
        with open('tokens.json') as f:
            tokens = json.load(f)
    else:
        tokens = None

    console.print(Panel.fit("Obtain Webex API Tokens", title="Step 1"))

    # Determine relevant route to obtain a valid Webex Token. Options include:
    # 1. Full OAuth (both primary and refresh token are expired)
    # 2. Simple refresh (only the primary token is expired)
    # 3. Valid token (the existing primary token is valid)
    if tokens is None or time.time() > (
            tokens['expires_at'] + (tokens['refresh_token_expires_in'] - tokens['expires_in'])):
        # Both tokens expired, run the OAuth Workflow
        console.print("[red]Both tokens are expired, we need to run OAuth workflow... See README.[/]")
        sys.exit(0)
    elif time.time() > tokens['expires_at']:
        # Generate a new token using the refresh token
        console.print("Existing primary token [red]expired[/]! Using refresh token...")

        tokens = refresh_token(tokens)
        generate_calling_report(tokens['access_token'])
    else:
        # Use existing valid token
        console.print("Existing primary token is [green]valid![/]")

        generate_calling_report(tokens['access_token'])


if __name__ == "__main__":
    main()
