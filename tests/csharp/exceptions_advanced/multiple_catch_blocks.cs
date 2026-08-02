// vybe-test: csharp/exceptions_advanced/multiple_catch_blocks
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

try {
    throw new ArgumentException("bad arg");
} catch (ArgumentNullException) {
    Console.WriteLine("null");
} catch (ArgumentException e) {
    Console.WriteLine("arg: " + e.Message);
} catch (Exception) {
    Console.WriteLine("generic");
}
