// vybe-test: csharp/csharp_string_builder/append_builds_string_incrementally
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

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

var sb = new System.Text.StringBuilder();
sb.Append("Hello"); sb.Append(" World");
__P((sb.ToString()).ToString());
__Check("Hello World");
