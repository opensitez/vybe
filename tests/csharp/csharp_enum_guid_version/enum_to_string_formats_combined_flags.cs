// vybe-test: csharp/csharp_enum_guid_version/enum_to_string_formats_combined_flags
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Permission { Read = 1, Write = 2 } var value = Permission.Read | Permission.Write; __Check((value.ToString()).ToString(), "Read, Write");
