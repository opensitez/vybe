// vybe-test: csharp/common_patterns/static_utility_class
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

static class StringUtils {
    public static bool IsPalindrome(string s) {
        string lower = s.ToLower();
        char[] chars = lower.ToCharArray();
        Array.Reverse(chars);
        return lower == new string(chars);
    }
}
__P((StringUtils.IsPalindrome("racecar")).ToString());
__P((StringUtils.IsPalindrome("hello")).ToString());
__P((StringUtils.IsPalindrome("Madam")).ToString());
__Check("True\nFalse\nTrue");
