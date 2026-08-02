// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
int? maybe = 97; __Check((maybe.HasValue && maybe.Value == 97).ToString(), "True");
