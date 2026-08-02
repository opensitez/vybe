// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
int? maybe = 32; __Check((maybe.HasValue && maybe.Value == 32).ToString(), "True");
