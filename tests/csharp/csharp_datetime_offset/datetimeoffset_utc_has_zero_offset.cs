// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_utc_has_zero_offset
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

using static __Harness;

var dto=System.DateTimeOffset.UtcNow;
__P((dto.Offset==System.TimeSpan.Zero).ToString());
__Check("True");

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
