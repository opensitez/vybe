// vybe-test: csharp/csharp_bcl_collections/read_only_list_view_reflects_backing_list_mutations
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var backing = new System.Collections.Generic.List<int> { 1 };
var view = backing.AsReadOnly();
backing.Add(2);
__Check((view.Count).ToString(), "2");
