// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_complement_within_byte_mask_inverts_bits
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
enum Perm : byte { A = 1, B = 2 }
var value = Perm.A | Perm.B;
var cleared = value & ~Perm.A;
__P(((int)cleared).ToString());
__Check("2");
