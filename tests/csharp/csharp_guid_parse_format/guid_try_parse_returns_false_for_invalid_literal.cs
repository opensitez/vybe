// vybe-test: csharp/csharp_guid_parse_format/guid_try_parse_returns_false_for_invalid_literal
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Guid value;
var ok = System.Guid.TryParse("not-a-guid", out value);
__Check((ok).ToString(), "False");
