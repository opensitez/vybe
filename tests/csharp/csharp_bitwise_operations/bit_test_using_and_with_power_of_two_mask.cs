// vybe-test: csharp/csharp_bitwise_operations/bit_test_using_and_with_power_of_two_mask
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

int flags = 0b1010; __P(((flags & 0b0010) != 0).ToString());
__Check("True");
