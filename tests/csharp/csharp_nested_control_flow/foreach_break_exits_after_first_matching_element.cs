// vybe-test: csharp/csharp_nested_control_flow/foreach_break_exits_after_first_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int hits = 0;
foreach (var value in new[] { 2, 4, 6, 8 }) {
    if (value == 6) break;
    hits++;
}
Console.WriteLine(hits);
