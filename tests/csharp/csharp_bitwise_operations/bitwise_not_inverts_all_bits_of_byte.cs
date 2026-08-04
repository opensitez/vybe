// vybe-test: csharp/csharp_bitwise_operations/bitwise_not_inverts_all_bits_of_byte
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operations.rs

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

byte b = 0b11110000; __P(((byte)(~b)).ToString());
__Check("15");
