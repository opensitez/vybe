// vybe-test: csharp/csharp_linq_query_syntax/linq_query_orders_words_by_length_then_name
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
var words = new[] { "pear", "fig", "banana", "kiwi" };
var query = from word in words
            orderby word.Length, word
            select word;
foreach (var word in query) __P((word).ToString());
__Check("fig\nkiwi\npear\nbanana");
