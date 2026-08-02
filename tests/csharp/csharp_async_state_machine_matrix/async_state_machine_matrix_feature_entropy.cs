// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
string feature = "async_state_machine_matrix:88"; __Check((feature.Length >= 1).ToString(), "True");
