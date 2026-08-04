// vybe-test: csharp/csharp_dictionary_operations/keys_collection_enumerates_all_inserted_keys
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

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

var d = new System.Collections.Generic.Dictionary<string,int>{{"x",1},{"y",2}};
var keys = new System.Collections.Generic.List<string>(d.Keys);
keys.Sort();
foreach(var k in keys) __P((k).ToString());
__Check("x\ny");
