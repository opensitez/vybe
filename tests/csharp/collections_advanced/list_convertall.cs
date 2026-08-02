// vybe-test: csharp/collections_advanced/list_convertall
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

var nums = new List<int> { 1, 2, 3 };
var strings = nums.ConvertAll(x => x.ToString());
foreach (var s in strings) Console.WriteLine(s);
