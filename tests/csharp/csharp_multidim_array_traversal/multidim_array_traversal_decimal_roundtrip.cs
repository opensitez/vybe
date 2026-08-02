// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
