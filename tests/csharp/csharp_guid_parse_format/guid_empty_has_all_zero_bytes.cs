// vybe-test: csharp/csharp_guid_parse_format/guid_empty_has_all_zero_bytes
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var id = System.Guid.Empty;
__Check((id == new System.Guid("00000000-0000-0000-0000-000000000000")).ToString(), "True");
