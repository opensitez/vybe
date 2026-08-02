// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_raw_string_embeds_variable
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count=7; string text=$"""items={count}"""; __Check((text).ToString(), "items=7");
