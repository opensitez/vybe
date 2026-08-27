// vybe-test: csharp/csharp_records_advanced/record_can_define_computed_property_from_primary_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

__P((new Rectangle(3, 7).Area).ToString());
__Check("21");

record Rectangle(int Width, int Height) { public int Area => Width * Height; }

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
