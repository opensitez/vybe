// vybe-test: csharp/csharp_type_aliases/using_alias_for_fully_qualified_type
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Dict=System.Collections.Generic.Dictionary<string,int>;
var d=new Dict{{"a",1},{"b",2}};
__Check((d["b"]).ToString(), "2");
