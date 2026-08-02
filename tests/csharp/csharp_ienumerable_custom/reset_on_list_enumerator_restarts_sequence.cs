// vybe-test: csharp/csharp_ienumerable_custom/reset_on_list_enumerator_restarts_sequence
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

var list=new System.Collections.Generic.List<int>{1,2,3};
int count=0;
foreach(var _ in list) count++;
foreach(var _ in list) count++;
Console.WriteLine(count);
