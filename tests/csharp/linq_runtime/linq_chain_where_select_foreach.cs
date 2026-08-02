// vybe-test: csharp/linq_runtime/linq_chain_where_select_foreach
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3); list.Add(4); list.Add(5); list.Add(6);
list.Where(x => x % 2 == 0).Select(x => x * 10).ForEach(x => Console.WriteLine(x));
