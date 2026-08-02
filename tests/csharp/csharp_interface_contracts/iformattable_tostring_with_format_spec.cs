// vybe-test: csharp/csharp_interface_contracts/iformattable_tostring_with_format_spec
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.IFormattable value = (object)3.14;
__Check((value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture)).ToString(), "3.1");
