// vybe-test: csharp/csharp_patterns/string_reversal
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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

class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}
__P((StringUtils.Reverse("hello")).ToString());
__P((StringUtils.Reverse("abcde")).ToString());
__Check("olleh\nedcba");
