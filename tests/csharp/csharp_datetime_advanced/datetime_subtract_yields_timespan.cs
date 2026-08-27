// vybe-test: csharp/csharp_datetime_advanced/datetime_subtract_yields_timespan
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

var a=new System.DateTime(2024,1,10);
var b=new System.DateTime(2024,1,1);
var diff=a-b;
__P((diff.Days).ToString());
__Check("9");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
