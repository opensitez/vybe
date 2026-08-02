// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
int? maybe = null; int fallback = maybe ?? 74; __Check((fallback == 74).ToString(), "True");
