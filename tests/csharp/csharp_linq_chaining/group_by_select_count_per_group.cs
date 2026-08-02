// vybe-test: csharp/csharp_linq_chaining/group_by_select_count_per_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

var words=new[]{"cat","car","bar","bat","can"};
var groups=words.GroupBy(w=>w[0])
    .Select(g=>(g.Key,g.Count()))
    .OrderBy(t=>t.Key);
foreach(var(k,c) in groups) Console.WriteLine($"{k}:{c}");
