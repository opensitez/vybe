// vybe-test: csharp/csharp_attributes_metadata/flags_attribute_marks_enum_but_bitwise_result_still_formats_value
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Flags] enum Permission { Read = 1, Write = 2, Execute = 4 } var permission = Permission.Read | Permission.Write; __Check((permission).ToString(), "Read, Write");
