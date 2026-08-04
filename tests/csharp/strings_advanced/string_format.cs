// vybe-test: csharp/strings_advanced/string_format
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

__P((string.Format("{0} + {1} = {2}", 1, 2, 3)).ToString());
__P((string.Format("Name: {0}, Age: {1}", "Bob", 25)).ToString());
__Check("1 + 2 = 3\nName: Bob, Age: 25");
