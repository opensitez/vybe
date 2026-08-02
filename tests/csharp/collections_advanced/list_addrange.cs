// vybe-test: csharp/collections_advanced/list_addrange
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var list = new List<int> { 1, 2, 3 };
list.AddRange(new int[] { 4, 5 });
Console.WriteLine(list.Count);
foreach (var x in list) Console.WriteLine(x);
