// vybe-test: csharp/csharp_list_operations/find_returns_first_element_satisfying_predicate
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{1,4,7,8};
__Check((list.Find(x => x > 5)).ToString(), "7");
