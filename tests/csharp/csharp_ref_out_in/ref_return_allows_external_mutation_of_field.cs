// vybe-test: csharp/csharp_ref_out_in/ref_return_allows_external_mutation_of_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

var g=new Grid();
g.Cell(1)=7;
__P((g.Get(1)).ToString());
__Check("7");

class Grid{
    int[] data={0,0,0};
    public ref int Cell(int i)=>ref data[i];
    public int Get(int i)=>data[i];
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
