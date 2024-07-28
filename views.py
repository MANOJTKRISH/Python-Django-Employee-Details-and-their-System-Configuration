from django.shortcuts import render
from django.http import JsonResponse
from .models import Employee
import platform
import psutil
import subprocess
import wmi
import socket

def fetch_employee_details(request):
    if request.method == 'POST':
        employee_number = request.POST.get('employee_number', None)
        if employee_number:
            try:
                employee = Employee.objects.get(Empid=employee_number)
                data = {
                    'name': employee.Name,
                    'cc_description': employee.CCDescpription,
                    'personal_area': employee.Personal_Area,
                    'department': employee.Personal_Sub_Area,
                    'designation': employee.Designation,
                    
                }
                return JsonResponse({'employee_details': data})
            except Employee.DoesNotExist:
                return JsonResponse({'error': 'Employee not found'}, status=404)
    return JsonResponse({'error': 'Invalid request'}, status=400)

def fetch_employee_form(request):
    
    processor = platform.processor()
    os_version = platform.platform()
    ram = f"{psutil.virtual_memory().total // (1024 ** 3)} GB" # RAM in GB
    hdd =f"{psutil.disk_usage('/').total // (1024 ** 3)} GB"
    system_model = platform.system()

    #print(processor)  # This will print to the server console
    #print(os_version) # This will print to the server console
    try:
        c = wmi.WMI()
        dvd_drive = c.CDROMDrive()[0]
        dvd_type = dvd_drive.MediaType if hasattr(dvd_drive, 'MediaType') else "Unknown"
    except Exception as e:
        dvd_type = "Not found"
    try:
        result = subprocess.run(['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version'],
                                capture_output=True, text=True, check=True)
        chrome_version = result.stdout.split()[-1]  # Extracting version number
    except subprocess.CalledProcessError:
        chrome_version = "Not found" 
    try:
        c = wmi.WMI()
        mouse_info = c.Win32_PointingDevice()[0]
        mouse_type = mouse_info.Description if hasattr(mouse_info, 'Description') else "Unknown"
    except Exception as e:
        mouse_type = "Not found"  
    try:
        c = wmi.WMI()
        computer_name = c.Win32_ComputerSystem()[0].Name
    except Exception as e:
        computer_name = "Not found"    
    try:
        c = wmi.WMI()
        # Query for Windows updates
        updates = c.Win32_QuickFixEngineering()
        patch_info = ", ".join(f"{update.HotFixID}: {update.Description}" for update in updates)
    except Exception as e:
        patch_info = "Not found" 
    
    try:
        # Query the Windows Registry for Microsoft Office version
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration'],
                                capture_output=True, text=True, check=True)
        
        office_version = None
        lines = result.stdout.splitlines()
        for line in lines:
            if line.strip().startswith("Version"):
                office_version = line.split()[2].strip()
                break
        
        if not office_version:
            office_version = "Not found"
    except Exception as e:
        office_version = "Not found" 
    try:
        # Execute 'java -version' command to get Java version
        result = subprocess.run(['java', '-version'], capture_output=True, text=True, check=True)
        
        # Extracting Java version from the output
        java_version = result.stderr.splitlines()[0].split()[2].strip('"')
    except Exception as e:
        java_version = "Not found"   
    try:
        # Query the Windows Registry for Adobe Reader version
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Adobe\\Acrobat Reader'], 
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            adobe_reader_version = version_line.split()[2].strip()
        else:
            adobe_reader_version = "Not found"
    except Exception as e:
        adobe_reader_version = "Not found"
    try:
        # Get the hostname
        hostname = socket.gethostname()
        # Get the IP address associated with the hostname
        ip_address = socket.gethostbyname(hostname)
    except Exception as e:
        ip_address = "Not found" 
        # Fetching SAP version
    try:
        # Example: querying a registry key for SAP version
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\SAP'], 
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            sap_version = version_line.split()[2].strip()
        else:
            sap_version = "Not found"
    except Exception as e:
        sap_version = "Not found"   # Fetching Outlook patch version
    try:
        # Example: querying a registry key for Outlook patch version
        result = subprocess.run(['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\Outlook\\Addins\\YourAddin', '/v', 'Version'],
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            outlook_patch_version = version_line.split()[2].strip()
        else:
            outlook_patch_version = "Not found"
    except Exception as e:
        outlook_patch_version = "Not found"
    # Fetching LAPS installed version
    try:
        # Example: querying a registry key for LAPS installed version
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\AdmPwd', '/v', 'Version'],
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            laps_installed_version = version_line.split()[2].strip()
        else:
            laps_installed_version = "Not found"
    except Exception as e:
        laps_installed_version = "Not found"    
    # Fetching system serial number using WMI
    try:
        c = wmi.WMI()
        system_serial_number = c.Win32_BIOS()[0].SerialNumber.strip()
    except Exception as e:
        system_serial_number = "Not found"    

    # Fetching system type using WMI
    try:
        c = wmi.WMI()
        system_type = c.Win32_ComputerSystem()[0].PCSystemType
        if system_type == 0:
            system_type = "Unspecified"
        elif system_type == 1:
            system_type = "Desktop"
        elif system_type == 2:
            system_type = "Mobile"
        elif system_type == 3:
            system_type = "Workstation"
        elif system_type == 4:
            system_type = "Enterprise Server"
        elif system_type == 5:
            system_type = "SOHO Server"
        elif system_type == 6:
            system_type = "Appliance PC"
        elif system_type == 7:
            system_type = "Performance Server"
        else:
            system_type = "Unknown"
    except Exception as e:
        system_type = "Not found" 
    # Fetching McAfee Agent version from registry
    try:
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\McAfee\\Agent', '/v', 'Version'],
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            mcafee_agent_version = version_line.split()[2].strip()
        else:
            mcafee_agent_version = "Not found"
    except Exception as e:
        mcafee_agent_version = "Not found" 
    # Fetching monitor model and size using WMI
    try:
        c = wmi.WMI()
        monitor = c.Win32_DesktopMonitor()[0]
        monitor_model = monitor.Description.strip()
        monitor_size = f"{monitor.ScreenWidth}x{monitor.ScreenHeight} pixels"
    except Exception as e:
        monitor_model = "Not found"
        monitor_size = "Not found" 
    # Fetching Symantec Management Agent version from registry
    try:
        result = subprocess.run(['reg', 'query', 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Symantec\\Symantec Agent', '/v', 'Version'],
                                capture_output=True, text=True, check=True)
        
        lines = result.stdout.splitlines()
        version_line = next((line for line in lines if line.strip().startswith("Version")), None)
        
        if version_line:
            symantec_management_version = version_line.split()[2].strip()
        else:
            symantec_management_version = "Not found"
    except Exception as e:
        symantec_management_version = "Not found"
                          
    # Return this data to a template
    context = {
        'processor': processor,
        'os_version': os_version,
        'ram': ram,
        'chrome_version': chrome_version,
        'dvd_type': dvd_type,
        'hdd': hdd,
        'system_model': system_model,
        'mouse_type': mouse_type,
        'computer_name': computer_name,
        'patch_info': patch_info,
        'office_version': office_version,
        'java_version': java_version,
        'adobe_reader_version': adobe_reader_version,
        'ip_address': ip_address,
        'sap_version': sap_version,
        'outlook_patch_version': outlook_patch_version,
        'laps_installed_version': laps_installed_version,
        'system_serial_number': system_serial_number,
        'system_type': system_type,
        'mcafee_agent_version': mcafee_agent_version,
        'monitor_model': monitor_model,
        'monitor_size': monitor_size,
        'symantec_management_version': symantec_management_version,

    }
    return render(request, 'index.html', context)

    #return render(request,"./index.html",{Processor:processor,Os_Ver:os_version})
