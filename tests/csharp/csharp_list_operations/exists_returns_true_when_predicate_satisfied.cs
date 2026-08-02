// vybe-test: csharp/csharp_list_operations/exists_returns_true_when_predicate_satisfied
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,2,3};
__Check((list.Exists(x => x > 2)).ToString(), "True");
