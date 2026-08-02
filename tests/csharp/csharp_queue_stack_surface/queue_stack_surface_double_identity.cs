// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
double seed = 32; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
