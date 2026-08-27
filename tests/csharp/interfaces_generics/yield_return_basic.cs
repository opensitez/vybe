// vybe-test: csharp/interfaces_generics/yield_return_basic
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

foreach (var n in Numbers.OneToFive()) __P((n).ToString());
__Check("1\n2\n3\n4\n5");

class Numbers {
    public static IEnumerable<int> OneToFive() {
        yield return 1;
        yield return 2;
        yield return 3;
        yield return 4;
        yield return 5;
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
