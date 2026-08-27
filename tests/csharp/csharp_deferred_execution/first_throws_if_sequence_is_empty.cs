// vybe-test: csharp/csharp_deferred_execution/first_throws_if_sequence_is_empty
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

using static __Harness;

string r="";
try{System.Array.Empty<int>().First();}
catch(System.InvalidOperationException){r="empty";}
__P((r).ToString());
__Check("empty");

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
