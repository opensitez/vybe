// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
int seed = 32; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
