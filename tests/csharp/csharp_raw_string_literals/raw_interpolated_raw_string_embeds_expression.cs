// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_raw_string_embeds_expression
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=2; int b=3; string text=$"""sum={a+b}"""; __Check((text).ToString(), "sum=5");
