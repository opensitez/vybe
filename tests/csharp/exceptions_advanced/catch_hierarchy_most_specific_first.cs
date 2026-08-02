// vybe-test: csharp/exceptions_advanced/catch_hierarchy_most_specific_first
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

try {
    throw new InvalidOperationException("invalid op");
} catch (InvalidOperationException e) {
    Console.WriteLine("specific: " + e.Message);
} catch (Exception) {
    Console.WriteLine("generic");
}
