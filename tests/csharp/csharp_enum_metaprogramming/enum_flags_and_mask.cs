// vybe-test: csharp/csharp_enum_metaprogramming/enum_flags_and_mask
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var v=(F.A|F.B|F.C)&F.B;
__P(((int)v).ToString());
__Check("2");

[System.Flags] enum F{A=1,B=2,C=4}

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
