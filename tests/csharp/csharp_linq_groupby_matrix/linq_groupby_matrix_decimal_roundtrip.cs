// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
