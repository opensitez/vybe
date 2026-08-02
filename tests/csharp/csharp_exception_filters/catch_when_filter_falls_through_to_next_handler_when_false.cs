// vybe-test: csharp/csharp_exception_filters/catch_when_filter_falls_through_to_next_handler_when_false
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

try {
    throw new System.Exception("code=500");
} catch (System.Exception e) when (e.Message.Contains("404")) {
    Console.WriteLine("not found");
} catch (System.Exception) {
    Console.WriteLine("server error");
}
