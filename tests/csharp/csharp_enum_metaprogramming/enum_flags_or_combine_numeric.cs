// vybe-test: csharp/csharp_enum_metaprogramming/enum_flags_or_combine_numeric
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

[System.Flags] enum F{A=1,B=2,C=4} var v=F.A|F.C; __P(((int)v).ToString());
__Check("5");
