// vybe-test: csharp/csharp_generics/generic_list_usage
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int>();
list.Add(10);
list.Add(20);
list.Add(30);
__Check((list.Count).ToString(), "3");
__Check((list[1]).ToString(), "20");
