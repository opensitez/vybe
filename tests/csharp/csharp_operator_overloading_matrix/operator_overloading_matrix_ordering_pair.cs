// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
int seed = 105; int right = seed + 1; __Check((seed < right).ToString(), "True");
