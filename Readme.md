﻿This script will add a URL to a URL Set. This version must be run locally on the TMG Server running the Firewall Service.

Description
Adds a url to an existing URL Set in ISA or TMG  

Functionality: 
Validates that it has received 2 parameters
Normalizes the Parameters
Connects to a TMG server
Looks for the looks for a URLSet indicated as first parameter
Attempts to add a URL indicated as second parameter
Attempts to add the subdomains for the URL indicated as second parameter
If any of the URL already exists the tool will reports it (it will not be duplicated in the URL Set)
Caveats
This Version must be run on a TMG server with the Firewall service running. (I'm working on another version w/o this restriction)
