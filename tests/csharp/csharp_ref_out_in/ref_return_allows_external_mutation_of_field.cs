// vybe-test: csharp/csharp_ref_out_in/ref_return_allows_external_mutation_of_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

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

class Grid{
    int[] data={0,0,0};
    public ref int Cell(int i)=>ref data[i];
    public int Get(int i)=>data[i];
}
var g=new Grid();
g.Cell(1)=7;
__P((g.Get(1)).ToString());
__Check("7");
