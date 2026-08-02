// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_in_foreach_loop_over_tuple_array
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

var pairs = new[]{(1,"a"),(2,"b"),(3,"c")};
int sum=0;
foreach(var (n, _) in pairs) sum+=n;
Console.WriteLine(sum);
