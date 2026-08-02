// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
var set = new System.Collections.Generic.HashSet<int>(); set.Add(71); set.Add(71); __Check((set.Count == 1).ToString(), "True");
