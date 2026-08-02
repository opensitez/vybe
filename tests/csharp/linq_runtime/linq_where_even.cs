// vybe-test: csharp/linq_runtime/linq_where_even
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5);
var evens = list.Where(x => x % 2 == 0);
evens.ForEach(x => Console.WriteLine(x));
