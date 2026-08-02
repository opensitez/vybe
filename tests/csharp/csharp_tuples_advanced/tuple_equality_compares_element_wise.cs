// vybe-test: csharp/csharp_tuples_advanced/tuple_equality_compares_element_wise
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = (1, "x"); var b = (1, "x");
__Check((a == b).ToString(), "True");
