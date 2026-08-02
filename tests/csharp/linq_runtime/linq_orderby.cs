// vybe-test: csharp/linq_runtime/linq_orderby
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(5); list.Add(3); list.Add(1); list.Add(4); list.Add(2);
list.OrderBy(x => x).ForEach(x => Console.WriteLine(x));
