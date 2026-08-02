// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
string feature = "queue_stack_surface:32"; __Check((feature.Length >= 1).ToString(), "True");
