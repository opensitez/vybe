// vybe-test: csharp/csharp_modern/using_static_class
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

__P((StringUtils.Reverse("hello")).ToString());
__Check("olleh");

static class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
