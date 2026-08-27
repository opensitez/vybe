// vybe-test: csharp/csharp_linq_query_syntax/linq_query_projects_anonymous_type_fields
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using static __Harness;
using System.Linq;

var prices = new[] { 3, 7, 10 }
;
var query = from price in prices
            select new TaggedTotal { Label = "item", Total = price * 2 }
;
foreach (var item in query) __P((item.Label + ":" + item.Total).ToString());
__Check("item:6\nitem:14\nitem:20");

class TaggedTotal {
    public string Label { get; set; }
    public int Total { get; set; }
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
