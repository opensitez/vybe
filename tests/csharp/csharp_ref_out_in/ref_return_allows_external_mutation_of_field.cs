// vybe-test: csharp/csharp_ref_out_in/ref_return_allows_external_mutation_of_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Grid{
    int[] data={0,0,0};
    public ref int Cell(int i)=>ref data[i];
    public int Get(int i)=>data[i];
}
var g=new Grid();
g.Cell(1)=7;
__Check((g.Get(1)).ToString(), "7");
