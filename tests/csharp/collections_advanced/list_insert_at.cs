// vybe-test: csharp/collections_advanced/list_insert_at
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var list = new List<string> { "a", "c", "d" };
list.Insert(1, "b");
foreach (var s in list) Console.WriteLine(s);
