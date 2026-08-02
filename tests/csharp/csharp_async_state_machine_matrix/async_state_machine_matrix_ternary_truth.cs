// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
int seed = 88; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
