// vybe-test: csharp/csharp_indexers/string_indexed_dictionary_wrapper_exposes_indexer
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Config{
    System.Collections.Generic.Dictionary<string,string> _d=new();
    public string this[string k]{get=>_d[k]; set=>_d[k]=value;}
}
var c=new Config(); c["env"]="prod";
__Check((c["env"]).ToString(), "prod");
