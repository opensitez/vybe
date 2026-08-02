// vybe-test: csharp/collections_advanced/list_toarray
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new List<int> { 10, 20, 30 };
int[] arr = list.ToArray();
__Check((arr.Length).ToString(), "3");
__Check((arr[1]).ToString(), "20");
