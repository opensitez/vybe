// vybe-test: csharp/csharp_parse_with_invariant_culture/int_parse_invariant_ignores_group_separators_in_strict_mode_failure
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

try {
    int.Parse("1,234", System.Globalization.CultureInfo.InvariantCulture);
    Console.WriteLine("parsed");
} catch (System.FormatException) {
    Console.WriteLine("reject");
}
