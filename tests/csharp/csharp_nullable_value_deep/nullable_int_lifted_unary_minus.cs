// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_lifted_unary_minus
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=8; __Check((-n).ToString(), "-8");
