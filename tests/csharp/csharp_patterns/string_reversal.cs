// vybe-test: csharp/csharp_patterns/string_reversal
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}
Console.WriteLine(StringUtils.Reverse("hello"));
Console.WriteLine(StringUtils.Reverse("abcde"));
