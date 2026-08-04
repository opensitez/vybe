// vybe-test: csharp/csharp_console_write/write_integer_then_write_integer_concatenate
// origin: languages/csharp/tests/csharp/test_csharp_console_write.rs

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

// console_write
__Pr((1).ToString()); __Pr((2).ToString()); __Pr((3).ToString()); __P("");
__Check("123");
