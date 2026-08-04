// vybe-test: csharp/csharp_collection_types/sorted_dictionary_first_key_is_smallest
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

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

var sd=new System.Collections.Generic.SortedDictionary<int,string>{{3,"c"},{1,"a"},{2,"b"}};
__P((sd.Keys.First()).ToString());
__Check("1");
