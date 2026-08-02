// vybe-test: csharp/exceptions_advanced/nested_try_inner_uncaught_propagates
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

try {
    try {
        throw new InvalidOperationException("oops");
    } catch (ArgumentException) {
        Console.WriteLine("wrong handler");
    }
} catch (InvalidOperationException e) {
    Console.WriteLine("outer got: " + e.Message);
}
