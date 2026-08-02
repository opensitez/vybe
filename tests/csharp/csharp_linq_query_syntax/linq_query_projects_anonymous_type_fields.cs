// vybe-test: csharp/csharp_linq_query_syntax/linq_query_projects_anonymous_type_fields
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
class TaggedTotal {
    public string Label { get; set; }
    public int Total { get; set; }
}
var prices = new[] { 3, 7, 10 };
var query = from price in prices
            select new TaggedTotal { Label = "item", Total = price * 2 };
foreach (var item in query) Console.WriteLine(item.Label + ":" + item.Total);
