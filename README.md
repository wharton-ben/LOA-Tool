# LOA-Tool
This project was designed to quickly remove a user from or set a user to leave of absence status.

## IMPORTANT NOTE
I have removed the specific names of the logon workstations. The Set-LOAUser funciton will set a limited set of logon workstations that the user can use during their leave. Usually this will be set to mail servers (so they can check their email). Once a specific logon workstation is set, all others will be prohibited. 
