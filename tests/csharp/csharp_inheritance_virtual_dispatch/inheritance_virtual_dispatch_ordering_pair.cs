// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
int seed = 71; int right = seed + 1; __Check((seed < right).ToString(), "True");
