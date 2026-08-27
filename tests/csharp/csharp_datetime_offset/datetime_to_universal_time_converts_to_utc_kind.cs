// vybe-test: csharp/csharp_datetime_offset/datetime_to_universal_time_converts_to_utc_kind
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

using static __Harness;

var local=new System.DateTime(2024,1,15,12,0,0,System.DateTimeKind.Local);
var utc=local.ToUniversalTime();
__P((utc.Kind).ToString());
__Check("Utc");

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
