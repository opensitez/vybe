// vybe-test: csharp/csharp_boxing_unboxing_matrix/boxing_unboxing_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_boxing_unboxing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// boxing_unboxing_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
