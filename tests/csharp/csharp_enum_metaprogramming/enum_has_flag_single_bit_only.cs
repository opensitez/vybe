// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_single_bit_only
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var v=Bit.Two;
__P((v.HasFlag(Bit.One)).ToString());
__P((v.HasFlag(Bit.Two)).ToString());
__Check("False\nTrue");

[System.Flags] enum Bit{One=1,Two=2}

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
