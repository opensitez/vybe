// vybe-test: csharp/csharp_closures/closure_sees_mutation_of_captured_variable
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

using static __Harness;

int x = 0;
System.Func<int> read = () => x;
x = 99;
__P((read()).ToString());
__Check("99");

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
