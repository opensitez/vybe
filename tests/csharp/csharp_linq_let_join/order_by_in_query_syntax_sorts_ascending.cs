// vybe-test: csharp/csharp_linq_let_join/order_by_in_query_syntax_sorts_ascending
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

var q=from n in new[]{3,1,2} orderby n select n;
foreach(var x in q) Console.WriteLine(x);
