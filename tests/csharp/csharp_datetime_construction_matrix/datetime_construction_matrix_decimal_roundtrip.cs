// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
