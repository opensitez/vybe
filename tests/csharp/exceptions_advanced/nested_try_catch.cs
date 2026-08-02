// vybe-test: csharp/exceptions_advanced/nested_try_catch
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

try {
    Console.WriteLine("outer try");
    try {
        throw new Exception("inner error");
    } catch (Exception e) {
        Console.WriteLine("inner catch: " + e.Message);
    }
    Console.WriteLine("after inner");
} catch (Exception) {
    Console.WriteLine("outer catch");
}
