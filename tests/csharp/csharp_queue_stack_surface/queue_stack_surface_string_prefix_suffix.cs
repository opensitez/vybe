// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
string feature = "queue_stack_surface"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
