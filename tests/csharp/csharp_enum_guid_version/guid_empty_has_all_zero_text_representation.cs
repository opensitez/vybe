// vybe-test: csharp/csharp_enum_guid_version/guid_empty_has_all_zero_text_representation
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Guid.Empty.ToString()).ToString(), "00000000-0000-0000-0000-000000000000");
