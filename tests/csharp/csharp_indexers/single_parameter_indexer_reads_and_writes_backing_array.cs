// vybe-test: csharp/csharp_indexers/single_parameter_indexer_reads_and_writes_backing_array
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

class Vec{
    int[] data=new int[3];
    public int this[int i]{get=>data[i]; set=>data[i]=value;}
}
var v=new Vec(); v[1]=42;
__P((v[1]).ToString());
__Check("42");
