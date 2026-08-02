// vybe-test: csharp/csharp_ienumerable_custom/manual_enumerator_move_next_and_current
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=new System.Collections.Generic.List<int>{10,20,30};
using var e=list.GetEnumerator();
e.MoveNext();
__Check((e.Current).ToString(), "10");
