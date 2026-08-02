// vybe-test: csharp/csharp_integer_arithmetic/addition_after_variable_assignment_accumulates_total
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int left = 9; int right = 11; __Check((left + right).ToString(), "20");
