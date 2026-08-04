// vybe-test: csharp/csharp_file_io/file_exists_returns_false_for_nonexistent_path
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

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

__P((System.IO.File.Exists("/no/such/path/xyz123.txt")).ToString());
__Check("False");
