// vybe-test: csharp/csharp_linq_projections/to_list_materializes_query_to_mutable_list
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new[]{1,2,3}.Select(x => x*2).ToList();
__Check((list.GetType().Name).ToString(), "List`1");
