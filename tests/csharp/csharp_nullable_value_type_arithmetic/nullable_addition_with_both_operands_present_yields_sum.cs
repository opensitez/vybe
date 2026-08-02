// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_addition_with_both_operands_present_yields_sum
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? left = 4;
int? right = 6;
int? sum = left + right;
__Check((sum.HasValue).ToString(), "True");
__Check((sum.Value).ToString(), "10");
