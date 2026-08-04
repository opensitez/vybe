// vybe-test: csharp/strings_advanced/stringbuilder_insert_replace
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

var sb = new System.Text.StringBuilder("Hello World");
sb.Replace("World", "There");
__P((sb.ToString()).ToString());
sb.Insert(5, " Beautiful");
__P((sb.ToString()).ToString());
__Check("Hello There\nHello Beautiful There");
