// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
int seed = 94; __Check((seed + 1 > seed).ToString(), "True");
