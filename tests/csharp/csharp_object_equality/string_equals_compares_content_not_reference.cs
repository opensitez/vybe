// vybe-test: csharp/csharp_object_equality/string_equals_compares_content_not_reference
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

using static __Harness;

string a = new string(new char[] { 'h', 'i' });
string b = new string(new char[] { 'h', 'i' });
__P((a.Equals(b)).ToString());
__Check("True");

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
