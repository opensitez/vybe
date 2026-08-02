// vybe-test: csharp/csharp_indexers/single_parameter_indexer_reads_and_writes_backing_array
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Vec{
    int[] data=new int[3];
    public int this[int i]{get=>data[i]; set=>data[i]=value;}
}
var v=new Vec(); v[1]=42;
__Check((v[1]).ToString(), "42");
