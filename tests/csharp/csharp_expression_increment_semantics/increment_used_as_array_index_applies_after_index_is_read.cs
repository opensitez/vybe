// vybe-test: csharp/csharp_expression_increment_semantics/increment_used_as_array_index_applies_after_index_is_read
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data = new[] { 10, 20, 30 };
int i = 0;
__Check((data[i++]).ToString(), "10");
__Check((i).ToString(), "1");
