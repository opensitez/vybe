// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// queue_stack_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[32] = 33; __Check((map.ContainsKey(32) && map[32] == 33).ToString(), "True");
