// vybe-test: csharp/csharp_records_advanced/record_with_mutable_property_can_be_updated_after_construction
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var box = new Box { Value = 3 }
;
box.Value = 8;
__P((box.Value).ToString());
__Check("8");

record Box { public int Value { get; set; } }

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
