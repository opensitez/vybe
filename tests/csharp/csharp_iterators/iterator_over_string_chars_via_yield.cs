// vybe-test: csharp/csharp_iterators/iterator_over_string_chars_via_yield
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

System.Collections.Generic.IEnumerable<char> Vowels(string s) {
    foreach(char c in s) if("aeiou".Contains(c)) yield return c;
}
int count=0;
foreach(var _ in Vowels("hello world")) count++;
__P((count).ToString());
__Check("3");

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
