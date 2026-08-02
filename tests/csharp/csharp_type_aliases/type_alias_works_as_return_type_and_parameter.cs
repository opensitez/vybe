// vybe-test: csharp/csharp_type_aliases/type_alias_works_as_return_type_and_parameter
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using NameMap=System.Collections.Generic.Dictionary<string,string>;
NameMap Build()=>new NameMap{{"k","v"}};
__Check((Build()["k"]).ToString(), "v");
