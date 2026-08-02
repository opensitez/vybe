// vybe-test: csharp/common_patterns/out_parameter
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((Parser.TryParse("42", out val)).ToString(), "True");
__Check((val).ToString(), "42");
__Check((Parser.TryParse("bad", out val)).ToString(), "False");
__Check((val).ToString(), "0");
