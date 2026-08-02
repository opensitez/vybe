// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
int seed = 94; int right = seed + 1; __Check((seed < right).ToString(), "True");
