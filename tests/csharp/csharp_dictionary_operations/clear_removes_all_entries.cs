// vybe-test: csharp/csharp_dictionary_operations/clear_removes_all_entries
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
d.Clear();
__Check((d.Count).ToString(), "0");
