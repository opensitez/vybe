// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
int seed = 71; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
