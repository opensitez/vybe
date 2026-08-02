// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
var map = new System.Collections.Generic.Dictionary<int, int>(); map[71] = 72; __Check((map.ContainsKey(71) && map[71] == 72).ToString(), "True");
