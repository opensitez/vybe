// vybe-test: csharp/csharp_linq_query_syntax/linq_query_uses_group_into_then_filters_large_groups
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var words = new[] { "ape", "ant", "boat", "berry", "cat" };
var query = from word in words
            group word by word.Length into groups
            where groups.Count() >= 2
            orderby groups.Key
            select groups.Key + ":" + groups.Count();
foreach (var item in query) Console.WriteLine(item);
