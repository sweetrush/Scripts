#!/bin/bash 


#Setting vm to not boot on XenServer Bootup 
xe vm-param-set uuid="7c396c6d-ef25-7b32-499d-5d3705f813b9" other-config:auto_poweron=true
xe vm-param-set uuid="1a3ef8a5-451f-dded-16bf-834c4ab199ff" other-config:auto_poweron=true
xe vm-param-set uuid="fe08c5e6-02d0-8323-9c08-9292a97c5d4c" other-config:auto_poweron=true
