// vybe-test: csharp/csharp_linq_quantifiers_partition/chunk_then_all_batches_full_except_last
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

var batches=new[]{1,2,3,4,5,6,7}.Chunk(3);
__P((batches.Take(2).All(b=>b.Length==3)?1:0).ToString());
__P((batches.Last().Length).ToString());
__Check("1\n1");
