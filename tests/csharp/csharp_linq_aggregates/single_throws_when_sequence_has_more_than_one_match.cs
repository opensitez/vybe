// vybe-test: csharp/csharp_linq_aggregates/single_throws_when_sequence_has_more_than_one_match
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

using static __Harness;

string result = "ok";
try { new[]{1,2}.Single(); }
catch(System.InvalidOperationException) { result = "many"; }
__P((result).ToString());
__Check("many");

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
