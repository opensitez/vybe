// vybe-test: csharp/csharp_linq_query_syntax/linq_query_combines_where_orderby_and_select
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
var values = new[] { 9, 1, 6, 2, 3 };
var query = from value in values
            where value >= 3
            orderby value descending
            select value - 1;
foreach (var item in query) __P((item).ToString());
__Check("8\n5\n2");
