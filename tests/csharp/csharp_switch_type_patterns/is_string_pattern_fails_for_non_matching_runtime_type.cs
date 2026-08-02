// vybe-test: csharp/csharp_switch_type_patterns/is_string_pattern_fails_for_non_matching_runtime_type
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

object boxed = 12;
if (boxed is string text) {
    Console.WriteLine(text);
} else {
    Console.WriteLine("not-string");
}
