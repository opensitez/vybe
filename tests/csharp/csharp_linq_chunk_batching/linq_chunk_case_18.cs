// vybe-test: csharp/csharp_linq_chunk_batching/linq_chunk_case_18

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int[] items = new int[] { 1, 2, 3, 4, 5, 6 };
var chunks = System.Linq.Enumerable.Chunk(items, 2).ToList();
__P(chunks.Count.ToString());
__P(chunks[0][0].ToString());
__Check("3\n1");
