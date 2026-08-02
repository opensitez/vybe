// vybe-test: csharp/csharp_collection_expressions/collection_expression_with_expression_elements
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] nums = [1 + 1, 2 + 2, 3 + 3];
__Check((nums[2]).ToString(), "6");
