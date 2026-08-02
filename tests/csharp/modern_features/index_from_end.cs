// vybe-test: csharp/modern_features/index_from_end
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] nums = { 10, 20, 30, 40, 50 };
__Check((nums[^1]).ToString(), "50");
__Check((nums[^2]).ToString(), "40");
