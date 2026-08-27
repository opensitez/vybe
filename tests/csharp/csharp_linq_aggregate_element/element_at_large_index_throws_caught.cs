// vybe-test: csharp/csharp_linq_aggregate_element/element_at_large_index_throws_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

string tag="ok";
try{new[]{1,2}.ElementAt(5);}
catch(System.ArgumentOutOfRangeException){tag="range";}
__P((tag).ToString());
__Check("range");

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
