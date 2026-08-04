// vybe-test: csharp/strings_advanced/string_insert_remove
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

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

string s = "Hello World";
__P((s.Insert(5, " Beautiful")).ToString());
__P((s.Remove(5)).ToString());
__P((s.Remove(5, 1)).ToString());
__Check("Hello Beautiful World\nHello\nHelloWorld");
