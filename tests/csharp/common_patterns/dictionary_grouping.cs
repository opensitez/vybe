// vybe-test: csharp/common_patterns/dictionary_grouping
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

var data = new List<string> { "apple", "banana", "avocado", "blueberry", "cherry" };
var grouped = new Dictionary<char, List<string>>();
foreach (var item in data) {
    char key = item[0];
    if (!grouped.ContainsKey(key)) grouped[key] = new List<string>();
    grouped[key].Add(item);
}
Console.WriteLine("a: " + grouped['a'].Count);
Console.WriteLine("b: " + grouped['b'].Count);
Console.WriteLine("c: " + grouped['c'].Count);
