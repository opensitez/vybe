// vybe-test: csharp/csharp_linq_chaining/chained_order_by_then_by_sorts_on_two_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

var data=new[]{(A:"b",B:2),(A:"a",B:3),(A:"a",B:1)};
var result=data.OrderBy(x=>x.A).ThenBy(x=>x.B);
foreach(var(a,b) in result) Console.WriteLine($"{a}{b}");
