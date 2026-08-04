// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_group_into_then_filters_large_groups
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
var words = new[] { "ape", "ant", "boat", "berry", "cat" };
var query = from word in words
            group word by word.Length into groups
            where groups.Count() >= 2
            orderby groups.Key
            select groups.Key + ":" + groups.Count();
foreach (var item in query) __P((item).ToString());
__Check("3:3");
