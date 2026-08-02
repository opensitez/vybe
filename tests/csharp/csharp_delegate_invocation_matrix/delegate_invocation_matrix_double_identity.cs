// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
double seed = 74; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
