// vybe-test: csharp/csharp_strings_ext/verbatim_string
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

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

var path = @"C:\Users\test\file.txt";
__P((path).ToString());
__Check("C:\\Users\\test\\file.txt");
