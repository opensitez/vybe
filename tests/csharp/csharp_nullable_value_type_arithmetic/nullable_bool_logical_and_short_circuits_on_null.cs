// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_bool_logical_and_short_circuits_on_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool? t = true;
bool? n = null;
bool? f = false;
__Check((t & n).ToString(), "");
__Check((n & f).ToString(), "");
__Check((f & t).ToString(), "False");
