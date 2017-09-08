# BizTalk Typed BAM API Generator
The GenerateTypedBamApi command line tool enables you to take a BizTalk BAM Observation model represented as a Excel Spreadsheet and generate a set of strongly typed C# classes which you can then use to create and populate BAM Activities.

The native BAM API is _loosely typed_ and therefore requires Activity Names and Activity Items to be supplied as string literals, this can be brittle especially as the observation model evolves over time and any typos, etc. will lead to runtime errors instead of compile time errors.

A strongly typed class is created for each Activity with properties for each Activity Item along with helper methods to write Activity Items to an Activity, Add References, Custom References (e.g. Message Bodies) and Continuation.

This tool uses a XSLT transform to turn the XML representing the BAM Observation Model into C# code.   

For further information on how to use BAM and why you might want to use a BAM API approach over the in-bult graphical Tracking Profile Editor please Chapter 6 of [Professional BizTalk Server 2006](http://www.wiley.com/WileyCDA/WileyTitle/productCd-0470046422.html).

## Example

Consider the code shown below which demonstates the creation of a BAM Activity (called Itinerary in this case), it populates a number of Activity Items, adds a reference to another Activity and finally adds a custom reference of a message body.  

As you can see from the code many string literals representing the Activity Name and Activity Items are required throughout the code, any mistakes will cause runtime errors which can be frustrating especially when you have to undeploy and redeploy your BizTalk solution.  Also, if you have a medium to large scale project then you will end up with many Activities and Activity Items which can require many lines of code to be manually created each time.

```c#
string ItineraryActivityID = System.Guid.NewGuid().ToString();
DirectEventStream des =
    new DirectEventStream("Integrated Security=SSPI;Data Source=.;Initial Catalog=BAMPrimaryImport", 1);
des.BeginActivity("Itinerary", ItineraryActivityID);

des.UpdateActivity("Itinerary", ItineraryActivityID, "Received", System.DateTime.Now,
    "Customer Name","Darren Jefford","County","Wiltshire","Total Itinerary Price",1285);
                
des.AddReference("Itinerary", ItineraryActivityID, "Activity", "Flight", flightActivityID);
des.AddReference("Itinerary", ItineraryActivityID, "MsgBody", "MessageBody", DateTime.Now.ToString(), myXmlMessageBody);

des.EndActivity("Itinerary",ItineraryActivityID);
```
In contrast consider the code shown below which uses the Interary C# class created by this tool, each Activity Item can be set using class properties which are strongly typed and therefore checked at compile time instead of runtime.  Simplified wrappers are provided around operations such as adding references and the ActivityID is stored internally once you've constructed the Itinerary class which saves you having to pass it each and every time.    In short this code is far simpler and easier to maintain especially as Activities evolve during your development lifecycle, also for medium to large projects this code generation approach can save you having to write hundreds to thousands of lines of code!

```c#
// Create a new Itinerary activity class, passing a GUID as the ActivityID
DirectESApi.Itinerary itin = new DirectESApi.Itinerary( System.Guid.NewGuid().ToString() );

// Begin the activity
itin.BeginItineraryActivity();

// Set activity items
itin.Received = System.DateTime.Now;
itin.CustomerName = "Darren Jefford";
itin.County = "Wiltshire";
itin.TotalItineraryPrice = 1285.00M;

// Commit these changes to the database;
itin.CommitItineraryActivity();

// Add a link between this Itinerary activity and another flight activity that already exists
itin.AddReferenceToAnotherActivity(DirectESApi.Activities.Flight, flightActivityID);

// Add a message body to this activity (or any other data you require)
itin.AddCustomReference("MsgBody", "MessageBody", DateTime.Now.ToString(), myXmlMessageBody);            

// End the Activity
itin.EndItineraryActivity();
```

## License

Copyright (c) Thomas F. Abraham. All rights reserved.

Licensed under the [MIT](LICENSE.txt) License.
