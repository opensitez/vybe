// vybe-test: csharp/csharp_linq_query_syntax/linq_query_groups_words_by_first_letter
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

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

using System.Linq;
var words = new[] { "apple", "ant", "banana", "boat" };
var groups = from word in words
             group word by word[0] into grouped
             orderby grouped.Key
             select grouped;
foreach (var group in groups) {
    __P((group.Key).ToString());
    __P((group.Count()).ToString());
}
__Check("a\n2\nb\n2");
