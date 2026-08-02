// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(32); set.Add(32); __Check((set.Count == 1).ToString(), "True");
