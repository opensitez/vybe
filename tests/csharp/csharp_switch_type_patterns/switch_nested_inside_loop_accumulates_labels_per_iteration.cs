// vybe-test: csharp/csharp_switch_type_patterns/switch_nested_inside_loop_accumulates_labels_per_iteration
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

string trace = "";
for (int i = 0; i < 3; i++) {
    trace += i switch { 0 => "a", 1 => "b", _ => "c" };
}
Console.WriteLine(trace);
