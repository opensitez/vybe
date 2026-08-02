// vybe-test: csharp/csharp_collections_initialise/collection_expression_creates_list_directly
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> list=[1,2,3];
__Check((list.Count).ToString(), "3"); __Check((list[1]).ToString(), "2");
