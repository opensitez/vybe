// vybe-test: csharp/csharp_type_aliases/using_alias_creates_shorter_name_for_type
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using IntList=System.Collections.Generic.List<int>;
var list=new IntList{1,2,3};
__Check((list.Count).ToString(), "3");
