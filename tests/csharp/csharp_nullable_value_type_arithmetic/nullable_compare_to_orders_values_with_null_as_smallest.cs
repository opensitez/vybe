// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_compare_to_orders_values_with_null_as_smallest
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? low = 1;
int? high = 5;
int? missing = null;
__Check((low.CompareTo(high)).ToString(), "-1");
__Check((missing.CompareTo(low)).ToString(), "-1");
