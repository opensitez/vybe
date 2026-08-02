// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
