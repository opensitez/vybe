// vybe-test: csharp/csharp_linq_complex/to_lookup_groups_items_by_key_returning_ilookup
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{(K:"a",V:1),(K:"a",V:2),(K:"b",V:3)};
var lu=data.ToLookup(x=>x.K,x=>x.V);
__Check((lu["a"].Sum()).ToString(), "3");
__Check((lu["b"].Sum()).ToString(), "3");
