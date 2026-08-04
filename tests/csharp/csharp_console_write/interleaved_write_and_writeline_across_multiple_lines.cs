// vybe-test: csharp/csharp_console_write/interleaved_write_and_writeline_across_multiple_lines
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

__Pr(("a").ToString()); __Pr(("b").ToString()); __P(("c").ToString()); __P(("d").ToString());
__Check("abc\nd");
