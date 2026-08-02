// vybe-test: csharp/csharp_enum_flags_operations/enum_parse_reads_name_into_typed_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green, Blue }
var value = System.Enum.Parse(typeof(Color), "Green");
__Check((value).ToString(), "Green");
