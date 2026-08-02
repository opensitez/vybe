// vybe-test: csharp/modern_features/var_inference
// origin: languages/csharp/tests/csharp/test_modern_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 42;
var s = "hello";
var list = new List<int> { 1, 2, 3 };
__Check((x.GetType().Name).ToString(), "Int32");
__Check((s.GetType().Name).ToString(), "String");
__Check((list.Count).ToString(), "3");
