Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to enable an OPC UA condition.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-ConnectivityStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-ConnectivityStudio-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://forum.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Define which server we will work with.
Dim endpointDescriptorUrlString
endpointDescriptorUrlString = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"

' Instantiate the client object.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

Dim methodNodeExpandedText
methodNodeExpandedText = "i=9027"   ' UAMethodIds.ConditionType_Enable

Dim objectNodeExpandedText
objectNodeExpandedText = "nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank?Yellow"
    
' Another way to create an empty array is to create an (empty) ElasticVector object, and convert it to an array.
    
Dim EmptyElasticVector: Set EmptyElasticVector = CreateObject("OpcLabs.BaseLib.Collections.ElasticVector")
Dim emptyArray: emptyArray = EmptyElasticVector.ToArray
    
' Perform the operation.
Client.CallMethod endpointDescriptorUrlString, objectNodeExpandedText, _
    methodNodeExpandedText, emptyArray, emptyArray

WScript.Echo "Done."

Rem#endregion Example
