// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_foreach
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

using static __Harness;

var item = new DeconstructItem("record_deconstruct_foreach", 42);
(string tag, int val) = item;
__P(tag);
__P(val.ToString());
__Check("record_deconstruct_foreach\n42");

record DeconstructItem(string Tag, int Val);
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
