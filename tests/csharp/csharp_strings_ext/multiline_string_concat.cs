// vybe-test: csharp/csharp_strings_ext/multiline_string_concat
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

string result = "Hello" +
    " " +
    "World";
__P((result).ToString());
__Check("Hello World");
