var scope = "ASCOM.OpenAstroTracker.Telescope";
var T = new ActiveXObject(scope);
T.Connected = true;

printInfo("Start Up");

slewRelativeMeridian(-0.00001, T.SiteLatitude);
printInfo("Post Slew");

T.SyncToCoordinates(5, -5);
printInfo("Post Sync");

WScript.StdOut.WriteLine("Script Complete");

function slewRelativeMeridian(degreesFromMeridian, degreesFromEquator){
    if( T.AtPark )
        T.Unpark();
    if (T.CanSetTracking && !T.Tracking)
        T.Tracking = true;
    var hoursFromMeridian = -degreesFromMeridian / 4;
    WScript.StdOut.WriteLine("Slewing to " + hoursFromMeridian + " from meridian...");
    T.SlewToCoordinates(T.SiderealTime + hoursFromMeridian, degreesFromEquator);
    WScript.StdOut.WriteLine("... slew complete");
}

function printInfo(title){
    WScript.StdOut.WriteLine("---------------------------- " + title);
    WScript.StdOut.WriteLine("RA:  " + T.RightAscension);
    WScript.StdOut.WriteLine("Dec: " + T.Declination);
    WScript.StdOut.WriteLine("Pier Side: " + T.SideOfPier);
    WScript.StdOut.WriteLine();
}