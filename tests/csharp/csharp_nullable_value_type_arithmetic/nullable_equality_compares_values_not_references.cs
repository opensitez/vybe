// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_equality_compares_values_not_references
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a = 7;
int? b = 7;
__Check((a == b).ToString(), "True");
int? c = null;
__Check((a == c).ToString(), "False");
__Check((c == null).ToString(), "True");
