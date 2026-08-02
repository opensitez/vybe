// vybe-test: csharp/csharp_tuples_advanced/tuple_in_linq_select_creates_anonymous_projection
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

var items = new[]{"apple","kiwi","pear"};
var proj = items.Select(s => (Name: s, Len: s.Length));
foreach(var x in proj) Console.WriteLine(x.Len);
