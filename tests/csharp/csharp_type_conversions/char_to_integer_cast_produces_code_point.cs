// vybe-test: csharp/csharp_type_conversions/char_to_integer_cast_produces_code_point
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch = 'A'; __Check(((int)ch).ToString(), "65");
