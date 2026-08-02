// vybe-test: csharp/csharp_type_conversions/enum_can_be_cast_to_underlying_integer
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode { Off = 0, On = 5 } __Check(((int)Mode.On).ToString(), "5");
