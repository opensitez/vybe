// vybe-test: csharp/csharp_linq_advanced/chunk_splits_sequence_into_fixed_size_batches
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

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

var batches=new[]{1,2,3,4,5}.Chunk(2).ToList();
__P((batches.Count).ToString());
__P((batches[0].Length).ToString());
__Check("3\n2");
