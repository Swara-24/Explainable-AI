import csv
import wmi

def get_driver_info(hostname):
    try:
        c = wmi.WMI(hostname)
        drivers = c.Win32_PnPSignedDriver()
        driver_info = []
        
        for driver in drivers:
            driver_name = driver.Name
            driver_version = driver.DriverVersion
            driver_info.append({'DriverName': driver_name, 'DriverVersion': driver_version})
        
        return driver_info
    except Exception as e:
        print(f"Failed to retrieve driver information for {hostname}: {str(e)}")
        return []

def write_to_csv(data, output_file):
    try:
        with open(output_file, mode='w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=['Hostname', 'DriverName', 'DriverVersion'])
            writer.writeheader()
            for item in data:
                writer.writerow({'Hostname': item['Hostname'], 'DriverName': item['DriverName'], 'DriverVersion': item['DriverVersion']})
        print(f"Driver information saved to {output_file}")
    except Exception as e:
        print(f"Failed to write to CSV file: {str(e)}")

def process_hostnames(hostnames, output_file):
    all_driver_info = []
    
    for hostname in hostnames:
        driver_info = get_driver_info(hostname)
        all_driver_info.extend([{'Hostname': hostname, 'DriverName': info['DriverName'], 'DriverVersion': info['DriverVersion']} for info in driver_info])
    
    write_to_csv(all_driver_info, output_file)

# Single hostname
hostname = 'WKST00012FFE.tus.ams1907.com'
output_file = 'driver_info.csv'
process_hostnames([hostname], output_file)

# Multiple hostnames from a text file
hostnames_file = 'hostnames.txt'
output_file = 'driver_info.csv'

with open(hostnames_file) as file:
    hostnames = [line.strip() for line in file if line.strip()]

process_hostnames(hostnames, output_file)
