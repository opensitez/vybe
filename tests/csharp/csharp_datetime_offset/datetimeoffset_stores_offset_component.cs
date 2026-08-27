// vybe-test: csharp/csharp_datetime_offset/datetimeoffset_stores_offset_component
// origin: languages/csharp/tests/csharp/test_csharp_datetime_offset.rs

using static __Harness;

var dto=new System.DateTimeOffset(2024,1,15,10,0,0,System.TimeSpan.FromHours(5));
__P((dto.Offset.Hours).ToString());
__Check("5");

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
