// vybe-test: csharp/csharp_linq_query_syntax/linq_query_groups_words_by_first_letter
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using static __Harness;
using System.Linq;

var words = new[] { "apple", "ant", "banana", "boat" }
;
var groups = from word in words
             group word by word[0] into grouped
             orderby grouped.Key
             select grouped;
foreach (var group in groups) {
    __P((group.Key).ToString());
    __P((group.Count()).ToString());
}
__Check("a\n2\nb\n2");

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
