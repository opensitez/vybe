// vybe-test: csharp/csharp_linq_query_syntax/linq_query_orders_words_by_length_then_name
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using static __Harness;
using System.Linq;

var words = new[] { "pear", "fig", "banana", "kiwi" }
;
var query = from word in words
            orderby word.Length, word
            select word;
foreach (var word in query) __P((word).ToString());
__Check("fig\nkiwi\npear\nbanana");

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
