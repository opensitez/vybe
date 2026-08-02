// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
int? maybe = null; int fallback = maybe ?? 32; __Check((fallback == 32).ToString(), "True");
