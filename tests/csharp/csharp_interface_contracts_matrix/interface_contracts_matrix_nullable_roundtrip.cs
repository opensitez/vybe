// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interface_contracts_matrix
int? maybe = 73; __Check((maybe.HasValue && maybe.Value == 73).ToString(), "True");
