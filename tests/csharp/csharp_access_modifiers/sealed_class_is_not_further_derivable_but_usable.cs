// vybe-test: csharp/csharp_access_modifiers/sealed_class_is_not_further_derivable_but_usable
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

using static __Harness;

__P((new Final().Value).ToString());
__Check("99");

sealed class Final{public int Value=99;}

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
