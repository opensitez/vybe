// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
