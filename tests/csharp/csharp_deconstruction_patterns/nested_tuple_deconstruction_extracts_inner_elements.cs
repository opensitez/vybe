// vybe-test: csharp/csharp_deconstruction_patterns/nested_tuple_deconstruction_extracts_inner_elements
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ((a,b),(c,d)) = ((1,2),(3,4));
__Check((a+b+c+d).ToString(), "10");
