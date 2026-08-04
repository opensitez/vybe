// vybe-test: csharp/csharp_enum_operations/flags_enum_combined_value_has_expected_integer
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

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

[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}
__P(((int)(Perm.Read|Perm.Execute)).ToString());
__Check("5");
