// vybe-test: csharp/csharp_integer_arithmetic/compound_assignment_chain_updates_final_value
// origin: languages/csharp/tests/csharp/test_csharp_integer_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int value = 10; value += 5; value *= 2; value -= 8; __Check((value).ToString(), "22");
