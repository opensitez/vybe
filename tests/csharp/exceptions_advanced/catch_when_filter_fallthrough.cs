// vybe-test: csharp/exceptions_advanced/catch_when_filter_fallthrough
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

try {
    throw new Exception("error 99");
} catch (Exception e) when (e.Message.Contains("42")) {
    Console.WriteLine("should not match");
} catch (Exception e) {
    Console.WriteLine("fallthrough: " + e.Message);
}
