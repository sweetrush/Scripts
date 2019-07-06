#!/bin/bash 
echo 'simple script for adding UUID machines '
echo "loading UUID $1"
if [ $1 == '' ]
	then
   echo 'No UUID Provided'
elif [ $1 != '' ]
 then
echo "Processing PownOn OPtiosn for UUID: $1"
 xe pool-param-set uuid=$1 other-config:auto_poweron=true
echo 'Added New UUID to auto poweron'
fi

