// vybe-test: csharp/csharp_bitwise_operations/compound_bitwise_and_assign_updates_in_place
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

int x = 0b1111; x &= 0b0101; __P((x).ToString());
__Check("5");
