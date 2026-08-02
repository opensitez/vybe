// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
int? maybe = null; int fallback = maybe ?? 97; __Check((fallback == 97).ToString(), "True");
