// vybe-test: csharp/csharp_number_bases/long_hex_literal_covers_full_64_bit_range
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

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

__P((0x7FFFFFFFFFFFFFFFL==long.MaxValue).ToString());
__Check("True");
