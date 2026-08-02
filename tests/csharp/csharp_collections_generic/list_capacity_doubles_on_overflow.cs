// vybe-test: csharp/csharp_collections_generic/list_capacity_doubles_on_overflow
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

var list=new System.Collections.Generic.List<int>(4);
for(int i=0;i<8;i++) list.Add(i);
Console.WriteLine(list.Count); Console.WriteLine(list.Capacity>=8);
