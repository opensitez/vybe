// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
int? maybe = null; int fallback = maybe ?? 73; __Check((fallback == 73).ToString(), "True");
