// vybe-test: csharp/csharp_indexers/interface_defines_indexer_contract_implemented_by_class
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

interface IMap{string this[int k]{get;}}
class Map:IMap{string[] data={"zero","one","two"};public string this[int k]=>data[k];}
IMap m=new Map();
__P((m[2]).ToString());
__Check("two");
