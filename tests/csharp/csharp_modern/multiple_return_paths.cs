// vybe-test: csharp/csharp_modern/multiple_return_paths
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Classify(int x) {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
}
__Check((Classify(5)).ToString(), "positive");
__Check((Classify(-3)).ToString(), "negative");
__Check((Classify(0)).ToString(), "zero");
