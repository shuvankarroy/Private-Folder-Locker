import os
import pyAesCrypt

bufferSize = 64 * 1024
key_file = open("keyfile.txt", "r")
key = key_file.readline()[1:-2]
key_file.close()

dir_file = open("dirfile.txt", "r")
dir_name = dir_file.readline()[1:-2]
dir_file.close()

if(key == "abrakadabra"):
    os.remove("keyfile.txt")
    os.remove("dirfile.txt")
    os.makedirs(dir_name)
    
