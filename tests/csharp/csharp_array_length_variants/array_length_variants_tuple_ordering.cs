// vybe-test: csharp/csharp_array_length_variants/array_length_variants_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
var tuple = (left: 25, right: 26); __Check((tuple.left < tuple.right).ToString(), "True");
