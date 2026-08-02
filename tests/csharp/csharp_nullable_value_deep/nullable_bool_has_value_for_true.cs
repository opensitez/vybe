// vybe-test: csharp/csharp_nullable_value_deep/nullable_bool_has_value_for_true
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool? flag=true; __Check((flag.HasValue).ToString(), "True"); __Check((flag.Value).ToString(), "True");
