// vybe-test: csharp/csharp_numeric_checked_bitwise/right_shift_moves_bits_to_lower_positions
// origin: languages/csharp/tests/csharp/test_csharp_numeric_checked_bitwise.rs

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

__P((16 >> 3).ToString());
__Check("2");
