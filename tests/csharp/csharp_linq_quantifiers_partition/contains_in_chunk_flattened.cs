// vybe-test: csharp/csharp_linq_quantifiers_partition/contains_in_chunk_flattened
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

var flat=new[]{1,2,3,4,5}.Chunk(2).SelectMany(x=>x);
__P((flat.Contains(5)?1:0).ToString());
__Check("1");
