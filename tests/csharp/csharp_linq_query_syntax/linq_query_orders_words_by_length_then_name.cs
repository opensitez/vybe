// vybe-test: csharp/csharp_linq_query_syntax/linq_query_orders_words_by_length_then_name
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_syntax.rs

using System.Linq;
var words = new[] { "pear", "fig", "banana", "kiwi" };
var query = from word in words
            orderby word.Length, word
            select word;
foreach (var word in query) Console.WriteLine(word);
