// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
int? maybe = 114; __Check((maybe.HasValue && maybe.Value == 114).ToString(), "True");
