// vybe-test: csharp/csharp_console_write/write_in_a_loop_builds_one_line
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

for (int i = 0; i < 3; i++) { __Pr((i).ToString()); } __P("");
__Check("012");
