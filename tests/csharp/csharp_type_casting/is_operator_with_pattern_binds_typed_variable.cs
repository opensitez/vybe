// vybe-test: csharp/csharp_type_casting/is_operator_with_pattern_binds_typed_variable
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = "world";
if(o is string s) __Check((s.Length).ToString(), "5");
