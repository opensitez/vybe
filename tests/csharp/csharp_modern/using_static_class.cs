// vybe-test: csharp/csharp_modern/using_static_class
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

static class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}
Console.WriteLine(StringUtils.Reverse("hello"));
