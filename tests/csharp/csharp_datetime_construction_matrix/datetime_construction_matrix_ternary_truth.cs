// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
int seed = 94; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
