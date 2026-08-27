// vybe-test: csharp/csharp_yield_iterators_core/yield_return_restart_iterator_second_foreach
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Two(){yield return 10;yield return 20;}
int sum=0;
foreach(var n in Two()) sum+=n;
foreach(var n in Two()) sum+=n;
__P((sum).ToString());
__Check("60");

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
