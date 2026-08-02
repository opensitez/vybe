// vybe-test: csharp/csharp_tuples_advanced/tuple_with_eight_elements_uses_rest_field
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t = (1,2,3,4,5,6,7,8);
__Check((t.Item8).ToString(), "8");
