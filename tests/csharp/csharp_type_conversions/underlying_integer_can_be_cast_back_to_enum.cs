// vybe-test: csharp/csharp_type_conversions/underlying_integer_can_be_cast_back_to_enum
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode { Off = 0, On = 5 } var mode = (Mode)5; __Check((mode).ToString(), "On");
