// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
var tuple = (left: 29, right: 30); __Check((tuple.left < tuple.right).ToString(), "True");
