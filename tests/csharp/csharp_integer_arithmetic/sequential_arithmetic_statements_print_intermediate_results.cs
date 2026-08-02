// vybe-test: csharp/csharp_integer_arithmetic/sequential_arithmetic_statements_print_intermediate_results
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 2; value = value + 3; __Check((value).ToString(), "5"); value = value * 4; __Check((value).ToString(), "20");
