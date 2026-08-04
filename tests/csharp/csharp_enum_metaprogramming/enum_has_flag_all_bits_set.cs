// vybe-test: csharp/csharp_enum_metaprogramming/enum_has_flag_all_bits_set
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

[System.Flags] enum Perm{Read=1,Write=2,Execute=4} var p=Perm.Read|Perm.Write|Perm.Execute; __P((p.HasFlag(Perm.Execute)).ToString());
__Check("True");
