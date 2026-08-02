// vybe-test: csharp/collections_advanced/list_removeat
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var list = new List<int> { 10, 20, 30, 40 };
list.RemoveAt(1);
foreach (var x in list) Console.WriteLine(x);
