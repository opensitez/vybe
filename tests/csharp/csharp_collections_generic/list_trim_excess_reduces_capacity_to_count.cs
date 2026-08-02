// vybe-test: csharp/csharp_collections_generic/list_trim_excess_reduces_capacity_to_count
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=new System.Collections.Generic.List<int>(100);
list.Add(1); list.Add(2);
list.TrimExcess();
__Check((list.Capacity<=list.Count*1).ToString(), "True");
