// vybe-test: csharp/csharp_strings/string_toupper_tolower
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

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

__P(("hello".ToUpper()).ToString());
__P(("HELLO".ToLower()).ToString());
__Check("HELLO\nhello");
