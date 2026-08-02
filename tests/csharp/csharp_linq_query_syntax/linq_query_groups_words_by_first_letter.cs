// vybe-test: csharp/csharp_linq_query_syntax/linq_query_groups_words_by_first_letter
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var words = new[] { "apple", "ant", "banana", "boat" };
var groups = from word in words
             group word by word[0] into grouped
             orderby grouped.Key
             select grouped;
foreach (var group in groups) {
    Console.WriteLine(group.Key);
    Console.WriteLine(group.Count());
}
