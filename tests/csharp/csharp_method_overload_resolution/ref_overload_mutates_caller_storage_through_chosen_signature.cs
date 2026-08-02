// vybe-test: csharp/csharp_method_overload_resolution/ref_overload_mutates_caller_storage_through_chosen_signature
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

void Scale(int value) { Console.WriteLine("byval:" + value); }
void Scale(ref int value) { value = value * 2; }
int n = 5;
Scale(ref n);
Console.WriteLine("after:" + n);
