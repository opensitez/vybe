// vybe-test: csharp/csharp_numeric_checked_bitwise/unchecked_block_wraps_byte_overflow_without_throwing
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

unchecked { byte value = 255; value += 1; __P((value).ToString()); }
__Check("0");
