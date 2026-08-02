// vybe-test: csharp/csharp_error_handling/try_catch_no_error
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

try {
    Console.WriteLine("ok");
} catch (Exception e) {
    Console.WriteLine("error");
}
