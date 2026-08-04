// vybe-test: csharp/csharp_enum_flags_operations/flags_enum_and_masks_to_intersection_of_bits
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
enum Perm { A = 1, B = 2, C = 4 }
var combined = Perm.A | Perm.B | Perm.C;
var masked = combined & Perm.B;
__P(((int)masked).ToString());
__Check("2");
