// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
string feature = "queue_stack_surface"; __Check((feature.Length > 0).ToString(), "True");
