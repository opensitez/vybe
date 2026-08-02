// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
var values = new System.Collections.Generic.List<int> { 88, 89, 88 }; __Check((values.Count == 3).ToString(), "True");
