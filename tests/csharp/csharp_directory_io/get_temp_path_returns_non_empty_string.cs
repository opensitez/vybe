// vybe-test: csharp/csharp_directory_io/get_temp_path_returns_non_empty_string
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

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

__P((System.IO.Path.GetTempPath().Length > 0).ToString());
__Check("True");
