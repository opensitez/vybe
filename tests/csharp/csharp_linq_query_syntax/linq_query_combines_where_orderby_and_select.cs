// vybe-test: csharp/csharp_linq_query_syntax/linq_query_combines_where_orderby_and_select
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var values = new[] { 9, 1, 6, 2, 3 };
var query = from value in values
            where value >= 3
            orderby value descending
            select value - 1;
foreach (var item in query) Console.WriteLine(item);
