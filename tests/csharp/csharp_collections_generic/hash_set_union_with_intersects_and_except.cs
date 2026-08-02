// vybe-test: csharp/csharp_collections_generic/hash_set_union_with_intersects_and_except
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new System.Collections.Generic.HashSet<int>{1,2,3,4};
var b=new System.Collections.Generic.HashSet<int>{3,4,5,6};
a.IntersectWith(b);
__Check((a.Count).ToString(), "2"); __Check((a.Contains(3)).ToString(), "True");
