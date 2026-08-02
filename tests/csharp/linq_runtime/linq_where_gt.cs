// vybe-test: csharp/linq_runtime/linq_where_gt
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30); list.Add(40); list.Add(50);
list.Where(x => x > 25).ForEach(x => Console.WriteLine(x));
