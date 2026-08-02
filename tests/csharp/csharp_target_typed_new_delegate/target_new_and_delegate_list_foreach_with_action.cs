// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_and_delegate_list_foreach_with_action
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

System.Collections.Generic.List<int> nums = new() { 1, 2, 3 };
int sum = 0;
System.Action<int> acc = n => sum += n;
foreach (var n in nums) acc(n);
Console.WriteLine(sum);
