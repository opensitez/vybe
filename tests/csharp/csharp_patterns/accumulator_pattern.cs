// vybe-test: csharp/csharp_patterns/accumulator_pattern
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var items = new[] { 1, 2, 3, 4, 5 }
;
int sum = 0;
int product = 1;
foreach (var x in items) {
    sum += x;
    product *= x;
}
__P((sum).ToString());
__P((product).ToString());
__Check("15\n120");

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
