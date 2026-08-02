// vybe-test: csharp/csharp_integer_arithmetic/addition_result_stored_in_new_variable
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int sum = 14 + 6; __Check((sum).ToString(), "20");
