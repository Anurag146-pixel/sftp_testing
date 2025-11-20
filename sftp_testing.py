import paramiko

hostname = "106.51.108.71"
port = 22
username = "sftpuser"
password = "admin123"

local_file = "sftp_test.txt"
remote_path = "/upload/" + local_file   # inside the chroot
download_file_name = "download.txt"

try:
    transport = paramiko.Transport((hostname, port))
    transport.connect(username=username, password=password)

    sftp = paramiko.SFTPClient.from_transport(transport)
    print("Connected to SFTP server")

    # Upload
    sftp.put(local_file, "/upload/" + local_file)
    print("File uploaded successfully")

    # Download
    sftp.get("/upload/" + local_file, download_file_name)
    print("File downloaded successfully")

    sftp.close()
    transport.close()
    print("Connection closed")

except Exception as e:
    print("Error:", e)
