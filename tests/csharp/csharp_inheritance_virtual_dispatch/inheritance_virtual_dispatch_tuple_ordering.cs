// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
var tuple = (left: 71, right: 72); __Check((tuple.left < tuple.right).ToString(), "True");
