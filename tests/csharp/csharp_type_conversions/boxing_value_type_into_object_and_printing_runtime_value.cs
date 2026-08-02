// vybe-test: csharp/csharp_type_conversions/boxing_value_type_into_object_and_printing_runtime_value
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 11; object boxed = count; __Check((boxed).ToString(), "11");
