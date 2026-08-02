// vybe-test: csharp/csharp_icomparable_sorting/comparer_default_sorts_strings_lexicographically
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<string>{"banana","apple","cherry"};
list.Sort(System.StringComparer.Ordinal);
__Check((list[0]).ToString(), "apple");
