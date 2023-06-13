#Exchange 2019 Transport Agent to update message class

This is a sample that shows how the message class of a message can be updated from within an Exchange 2019 Transport Agent.

Some assumptions are made (can be changed in the code):
- A log file will be created in the folder **C:\TA**.
- Only email messages (IPM.Note) where the subject starts with **UPDATEMESSAGECLASS** will be updated (all other messages are ignored).
- The new message class applied is **IPM.Note.Custom**.

To install the transport agent (assuming the binaries are in the **C:\TA** directory):
```
Install-TransportAgent -Name "UpdateMessageClassAgent" -TransportAgentFactory TA_UpdateMessageClass.UpdateMessageClassAgentFactory -AssemblyPath "C:\TA\TA_UpdateMessageClass.dll"
Enable-TransportAgent UpdateMessageClassAgent
```
