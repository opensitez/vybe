// vybe-test: csharp/csharp_indexers/string_indexed_dictionary_wrapper_exposes_indexer
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Config{
    System.Collections.Generic.Dictionary<string,string> _d=new();
    public string this[string k]{get=>_d[k]; set=>_d[k]=value;}
}
var c=new Config(); c["env"]="prod";
__P((c["env"]).ToString());
__Check("prod");
