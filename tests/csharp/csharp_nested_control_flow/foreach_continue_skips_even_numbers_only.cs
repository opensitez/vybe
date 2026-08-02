// vybe-test: csharp/csharp_nested_control_flow/foreach_continue_skips_even_numbers_only
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int sum = 0;
foreach (var value in new[] { 1, 2, 3, 4, 5 }) {
    if (value % 2 == 0) continue;
    sum += value;
}
Console.WriteLine(sum);
