// vybe-test: csharp/interfaces_generics/yield_return_fibonacci
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

foreach (var n in Fib.Sequence(8)) __P((n).ToString());
__Check("0\n1\n1\n2\n3\n5\n8\n13");

class Fib {
    public static IEnumerable<int> Sequence(int count) {
        int a = 0, b = 1;
        for (int i = 0; i < count; i++) {
            yield return a;
            int temp = a + b;
            a = b;
            b = temp;
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
