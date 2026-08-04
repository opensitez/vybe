// vybe-test: csharp/csharp_with_expression_records/with_double_field
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

record Rate(double V); var s=(new Rate(1.1)) with{V=2.2}; __P((s.V).ToString());
__Check("2.2");
