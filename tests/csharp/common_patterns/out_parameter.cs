// vybe-test: csharp/common_patterns/out_parameter
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

class Parser {
    public static bool TryParse(string s, out int result) {
        if (s == "42") { result = 42; return true; }
        result = 0;
        return false;
    }
}
int val;
__P((Parser.TryParse("42", out val)).ToString());
__P((val).ToString());
__P((Parser.TryParse("bad", out val)).ToString());
__P((val).ToString());
__Check("True\n42\nFalse\n0");
