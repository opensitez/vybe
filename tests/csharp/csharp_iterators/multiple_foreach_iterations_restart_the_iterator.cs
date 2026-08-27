// vybe-test: csharp/csharp_iterators/multiple_foreach_iterations_restart_the_iterator
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

System.Collections.Generic.IEnumerable<int> Three() {
    yield return 1; yield return 2; yield return 3;
}
int total=0;
foreach(var x in Three()) total+=x;
foreach(var x in Three()) total+=x;
__P((total).ToString());
__Check("12");

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
