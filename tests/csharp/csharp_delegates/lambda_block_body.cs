// vybe-test: csharp/csharp_delegates/lambda_block_body
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Func<int, string> classify = x => {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
};
__Check((classify(5)).ToString(), "positive");
__Check((classify(-3)).ToString(), "negative");
__Check((classify(0)).ToString(), "zero");
