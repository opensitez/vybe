// vybe-test: csharp/csharp_tuple_patterns/nested_tuple_pattern_matches_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=((1,2),(3,4));
var((a,b),(c,d))=data;
__Check((a+b+c+d).ToString(), "10");
