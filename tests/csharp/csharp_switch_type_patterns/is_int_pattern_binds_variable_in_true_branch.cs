// vybe-test: csharp/csharp_switch_type_patterns/is_int_pattern_binds_variable_in_true_branch
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

object boxed = 12;
if (boxed is int value) {
    Console.WriteLine(value + 1);
} else {
    Console.WriteLine("no");
}
