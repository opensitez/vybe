// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_capacity_then_add
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

System.Collections.Generic.List<int> buf = new();
for (int i = 0; i < 4; i++) buf.Add(i);
Console.WriteLine(buf.Count);
