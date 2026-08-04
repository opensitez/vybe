// vybe-test: csharp/modern_features/switch_expression_type_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

object obj = 42;
string result = obj switch {
    int i => "int: " + i,
    string s => "string: " + s,
    _ => "unknown"
};
__P((result).ToString());
__Check("int: 42");
