import re
import textfsm


def get_show_version(logfile):

    # Load the input file to a variable
    input_file = open(logfile)
    raw_text_data = input_file.read()
    input_file.close()

    # Get SH VERSION Raw Data
    raw_version_data = re.search(r'#sh ver.*#sh mod', raw_text_data, re.DOTALL)

    # Run SHOW VERSION through the FSM.
    show_version_template = open("cisco_ios_show_version.template")
    show_version_table = textfsm.TextFSM(show_version_template)
    show_version_parsed = show_version_table.ParseText(raw_version_data.group(0))

    return show_version_table.header, show_version_parsed


def get_show_int_status(logfile):

    # Load the input file to a variable
    input_file = open(logfile)
    raw_text_data = input_file.read()
    input_file.close()

    # Get SH INT STAT Raw Data
    raw_interfaces_status_data = re.search(r'#sh int status.*#sh int trunk', raw_text_data, re.DOTALL)

    # Run SHOW INTERFACES STATUS through the FSM.
    show_interfaces_status_template = open("cisco_ios_show_interfaces_status.template")
    show_interfaces_status_table = textfsm.TextFSM(show_interfaces_status_template)
    show_interfaces_status_parsed = show_interfaces_status_table.ParseText(raw_interfaces_status_data.group(0))

    return show_interfaces_status_table.header, show_interfaces_status_parsed


def get_show_cdp_nei_det(logfile):

    # Load the input file to a variable
    input_file = open(logfile)
    raw_text_data = input_file.read()
    input_file.close()

    # Get SH CDP NEI Raw Data
    raw_cdp_nei_det_data = re.search(r'#sh cdp nei det.*#sh int desc', raw_text_data, re.DOTALL)

    # Run SHOW CDP NEI DET through the FSM.
    show_cdp_nei_det_template = open("cisco_ios_show_cdp_neighbors_detail.template")
    show_cdp_nei_det_table = textfsm.TextFSM(show_cdp_nei_det_template)
    show_cdp_nei_det_parsed = show_cdp_nei_det_table.ParseText(raw_cdp_nei_det_data.group(0))

    # Normalize Interface Name
    for line in show_cdp_nei_det_parsed:
        for i in range(0, len(line)):
            if 'GigabitEthernet' in line[i]:
                line[i] = re.sub('^GigabitEthernet', 'Gi', line[i])

            elif 'TenGigabitEthernet' in line[i]:
                line[i] = re.sub('^TenGigabitEthernet', 'Te', line[i])

    return show_cdp_nei_det_table.header, show_cdp_nei_det_parsed
