// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
int seed = 74; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
