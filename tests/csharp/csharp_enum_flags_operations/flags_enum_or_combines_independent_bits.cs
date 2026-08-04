// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_or_combines_independent_bits
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

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

[System.Flags]
enum Perm { None = 0, Read = 1, Write = 2 }
var value = Perm.Read | Perm.Write;
__P(((int)value).ToString());
__Check("3");
