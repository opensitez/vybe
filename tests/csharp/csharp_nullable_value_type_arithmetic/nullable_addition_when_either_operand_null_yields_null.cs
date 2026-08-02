// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_addition_when_either_operand_null_yields_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? present = 5;
int? missing = null;
int? sum = present + missing;
__Check((sum.HasValue).ToString(), "False");
