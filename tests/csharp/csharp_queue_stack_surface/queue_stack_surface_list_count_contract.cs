// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
var values = new System.Collections.Generic.List<int> { 32, 33, 32 }; __Check((values.Count == 3).ToString(), "True");
