// vybe-test: csharp/csharp_linq_complex/group_by_then_order_group_keys_ascending
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

var items=new[]{(Cat:"b",Val:2),(Cat:"a",Val:1),(Cat:"b",Val:4),(Cat:"a",Val:3)};
var groups=items.GroupBy(i=>i.Cat).OrderBy(g=>g.Key)
    .Select(g=>(g.Key,g.Sum(i=>i.Val)));
foreach(var(k,s) in groups) Console.WriteLine($"{k}:{s}");
