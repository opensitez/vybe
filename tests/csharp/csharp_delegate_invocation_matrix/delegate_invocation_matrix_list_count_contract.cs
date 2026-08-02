// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
var values = new System.Collections.Generic.List<int> { 74, 75, 74 }; __Check((values.Count == 3).ToString(), "True");
