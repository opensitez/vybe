// vybe-test: csharp/csharp_bitwise_operations/right_shift_divides_by_power_of_two
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

__P((64 >> 3).ToString());
__Check("8");
