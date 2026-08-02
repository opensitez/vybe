// vybe-test: csharp/common_patterns/static_utility_class
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((StringUtils.IsPalindrome("racecar")).ToString(), "True");
__Check((StringUtils.IsPalindrome("hello")).ToString(), "False");
__Check((StringUtils.IsPalindrome("Madam")).ToString(), "True");
