// vybe-test: csharp/csharp_generic_variance2/ienumerable_is_covariant_over_its_element_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<string> strings=new[]{"a","b"};
System.Collections.Generic.IEnumerable<object> objects=strings;
__Check((objects.Count()).ToString(), "2");
