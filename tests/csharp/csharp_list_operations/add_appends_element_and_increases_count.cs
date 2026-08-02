// vybe-test: csharp/csharp_list_operations/add_appends_element_and_increases_count
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>();
list.Add(1); list.Add(2);
__Check((list.Count).ToString(), "2");
