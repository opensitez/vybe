// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_sorted_set_inferred_if_available
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

System.Collections.Generic.SortedSet<int> ordered = new();
ordered.Add(3); ordered.Add(1); ordered.Add(2);
foreach (var n in ordered) Console.WriteLine(n);
