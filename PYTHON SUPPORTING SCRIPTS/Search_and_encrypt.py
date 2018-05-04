import os
import pyAesCrypt

# For encryption purpose
bufferSize = 64 * 1024

key_file = open("keyfile.txt", "r")
key = key_file.readline()[1:-2]
key_file.close()

dir_file = open("dirfile.txt", "r")
rootdir = dir_file.readline()[1:-2]
dir_file.close()

if(key == "abrakadabra"):
    os.remove("keyfile.txt")
    os.remove("dirfile.txt")
    for subdir, dirs, files in os.walk(rootdir):
        for file in files:
            filepath = subdir + os.sep + file
            pyAesCrypt.encryptFile(filepath, filepath + ".aes", key, bufferSize)
            os.remove(filepath)

    # For Hiding the secure folder
    from platform import system
    folderPath = rootdir
    from subprocess import call
    call(["attrib", "+H","+S", folderPath])
