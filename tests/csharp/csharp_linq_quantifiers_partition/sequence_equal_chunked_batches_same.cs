// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_chunked_batches_same
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

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

var a=new[]{1,2,3,4};
var b=new[]{1,2,3,4};
__P((a.Chunk(2).SelectMany(x=>x).SequenceEqual(b)).ToString());
__Check("True");
