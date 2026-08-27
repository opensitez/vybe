// vybe-test: csharp/csharp_linq_aggregate_element/single_with_predicate_many_matches_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

string tag="ok";
try{new[]{1,2,2}.Single(x=>x==2);}
catch(System.InvalidOperationException){tag="many";}
__P((tag).ToString());
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
