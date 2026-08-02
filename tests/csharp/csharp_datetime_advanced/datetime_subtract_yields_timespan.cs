// vybe-test: csharp/csharp_datetime_advanced/datetime_subtract_yields_timespan
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new System.DateTime(2024,1,10);
var b=new System.DateTime(2024,1,1);
var diff=a-b;
__Check((diff.Days).ToString(), "9");
