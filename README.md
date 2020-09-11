# admindownCheck
Router physical port security enhancement tool...   
Gather Physical port operation status  
Compatible Model: C35xx, C37xx, C38xx, C65XX, Nexus 5K, 7K, 9K  
Scripted by Ethan Park, Sep. 2020

## Main logic
  - Reqeust a device list to AKiPS  
  - Access to the jumpbox  
  - Access to the device  
  - Gather a interface list with "down" status from each device
  - Using this interface to retrieve running-configuration (sh run int xxx) and check it has shutdown configuration
  - Save as a xlsx file
