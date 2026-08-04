// vybe-test: csharp/csharp_with_expression_records/with_short_field
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

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

record ShortBox(short S); var t=(new ShortBox(1)) with{S=1000}; __P((t.S).ToString());
__Check("1000");
