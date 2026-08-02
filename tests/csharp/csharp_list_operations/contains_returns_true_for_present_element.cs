// vybe-test: csharp/csharp_list_operations/contains_returns_true_for_present_element
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{10,20,30};
__Check((list.Contains(20)).ToString(), "True");
