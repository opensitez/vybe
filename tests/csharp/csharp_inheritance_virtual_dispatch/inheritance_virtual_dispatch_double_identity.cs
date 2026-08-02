// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
double seed = 71; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
