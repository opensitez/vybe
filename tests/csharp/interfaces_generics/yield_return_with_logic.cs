// vybe-test: csharp/interfaces_generics/yield_return_with_logic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

foreach (var n in Gen.EvenNumbers(10)) __P((n).ToString());
__Check("0\n2\n4\n6\n8\n10");

class Gen {
    public static IEnumerable<int> EvenNumbers(int max) {
        for (int i = 0; i <= max; i++) {
            if (i % 2 == 0) yield return i;
        }
    }
}

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
