// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_string_with_conditional_expression
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=4; string text=$"""{ (n%2==0 ? "even" : "odd") }"""; __Check((text.Trim()).ToString(), "even");
