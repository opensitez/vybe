// vybe-test: csharp/csharp_immutable_collections/immutable_array_indexer_reads_element
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

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

var arr=System.Collections.Immutable.ImmutableArray.Create(10,20,30);
__P((arr[1]).ToString());
__Check("20");
