// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
