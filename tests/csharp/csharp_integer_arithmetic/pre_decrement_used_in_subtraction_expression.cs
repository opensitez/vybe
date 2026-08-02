// vybe-test: csharp/csharp_integer_arithmetic/pre_decrement_used_in_subtraction_expression
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 9; int prior = value; int now = --value; __Check((prior - now).ToString(), "1");
