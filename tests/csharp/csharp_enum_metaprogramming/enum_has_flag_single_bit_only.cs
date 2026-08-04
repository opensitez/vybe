// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_single_bit_only
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Bit{One=1,Two=2} var v=Bit.Two; __P((v.HasFlag(Bit.One)).ToString()); __P((v.HasFlag(Bit.Two)).ToString());
__Check("False\nTrue");
