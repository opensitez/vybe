// vybe-test: csharp/csharp_nested_control_flow/foreach_iteration_variable_is_fresh_each_iteration
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int last = -1;
foreach (var value in new[] { 1, 2, 3 }) {
    last = value;
}
Console.WriteLine(last);
