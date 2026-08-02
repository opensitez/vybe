// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_to_string_null_prints_empty
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? n=null; __Check((n.ToString().Length).ToString(), "0");
