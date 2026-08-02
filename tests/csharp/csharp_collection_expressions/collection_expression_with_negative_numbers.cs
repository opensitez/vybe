// vybe-test: csharp/csharp_collection_expressions/collection_expression_with_negative_numbers
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] nums = [-1, -2];
__Check((nums[0] + nums[1]).ToString(), "-3");
