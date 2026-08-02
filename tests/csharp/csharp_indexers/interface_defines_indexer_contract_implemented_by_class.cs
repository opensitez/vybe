// vybe-test: csharp/csharp_indexers/interface_defines_indexer_contract_implemented_by_class
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IMap{string this[int k]{get;}}
class Map:IMap{string[] data={"zero","one","two"};public string this[int k]=>data[k];}
IMap m=new Map();
__Check((m[2]).ToString(), "two");
