// vybe-test: csharp/csharp_collection_expressions/collection_expression_modifying_copy_not_source
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src = [1, 2];
int[] copy = [..src];
copy[0] = 9;
__Check((src[0]).ToString(), "1"); __Check((copy[0]).ToString(), "9");
