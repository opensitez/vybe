// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
int seed = 32; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
