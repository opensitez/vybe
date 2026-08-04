// vybe-test: csharp/type_features/range_slice_string
// origin: languages/csharp/tests/csharp/test_type_features.rs

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
        var sub = s[0..5];
        __P((sub).ToString());
__Check("Hello");
