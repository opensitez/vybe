// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
var values = new System.Collections.Generic.List<int> { 71, 72, 71 }; __Check((values.Count == 3).ToString(), "True");
