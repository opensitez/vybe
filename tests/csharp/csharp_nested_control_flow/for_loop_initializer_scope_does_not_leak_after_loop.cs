// vybe-test: csharp/csharp_nested_control_flow/for_loop_initializer_scope_does_not_leak_after_loop
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int total = 0;
for (int i = 0; i < 3; i++) total += i;
Console.WriteLine(total);
