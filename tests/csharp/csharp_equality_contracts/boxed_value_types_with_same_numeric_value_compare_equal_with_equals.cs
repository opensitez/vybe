// vybe-test: csharp/csharp_equality_contracts/boxed_value_types_with_same_numeric_value_compare_equal_with_equals
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object left = 42;
object right = 42;
__Check((left.Equals(right)).ToString(), "True");
