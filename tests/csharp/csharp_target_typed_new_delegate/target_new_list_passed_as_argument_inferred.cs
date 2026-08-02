// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_passed_as_argument_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

int Sum(System.Collections.Generic.List<int> xs) { int s = 0; foreach (var x in xs) s += x; return s; }
System.Collections.Generic.List<int> data = new() { 1, 2, 3 };
Console.WriteLine(Sum(data));
