// vybe-test: csharp/collections_advanced/list_findall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var list = new List<int> { 1, 2, 3, 4, 5, 6 };
var evens = list.FindAll(x => x % 2 == 0);
foreach (var x in evens) Console.WriteLine(x);
