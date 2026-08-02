// vybe-test: csharp/modern_features/tuple_equality
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var t1 = (1, 2);
var t2 = (1, 2);
var t3 = (1, 3);
__Check((t1 == t2).ToString(), "True");
__Check((t1 == t3).ToString(), "False");
