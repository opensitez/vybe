// vybe-test: csharp/csharp_tuples_advanced/named_tuple_elements_accessed_by_name
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = (X: 3, Y: 4);
__Check((p.X).ToString(), "3"); __Check((p.Y).ToString(), "4");
