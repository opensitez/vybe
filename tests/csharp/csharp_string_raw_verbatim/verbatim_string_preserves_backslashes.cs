// vybe-test: csharp/csharp_string_raw_verbatim/verbatim_string_preserves_backslashes
// origin: languages/csharp/tests/csharp/test_csharp_string_raw_verbatim.rs

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

string path=@"C:\Users\test\file.txt";
__P((path.Contains(@"\test")).ToString());
__Check("True");
