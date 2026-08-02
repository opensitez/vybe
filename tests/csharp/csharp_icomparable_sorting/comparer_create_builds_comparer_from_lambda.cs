// vybe-test: csharp/csharp_icomparable_sorting/comparer_create_builds_comparer_from_lambda
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var cmp = System.Collections.Generic.Comparer<string>.Create(
    (a,b) => a.Length.CompareTo(b.Length));
var list = new System.Collections.Generic.List<string>{"cc","aaa","b"};
list.Sort(cmp);
__Check((list[0]).ToString(), "b");
