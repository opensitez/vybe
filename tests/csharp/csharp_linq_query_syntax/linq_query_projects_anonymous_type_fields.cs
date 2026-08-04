// vybe-test: csharp/csharp_linq_query_syntax/linq_query_projects_anonymous_type_fields
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
class TaggedTotal {
    public string Label { get; set; }
    public int Total { get; set; }
}
var prices = new[] { 3, 7, 10 };
var query = from price in prices
            select new TaggedTotal { Label = "item", Total = price * 2 };
foreach (var item in query) __P((item.Label + ":" + item.Total).ToString());
__Check("item:6\nitem:14\nitem:20");
