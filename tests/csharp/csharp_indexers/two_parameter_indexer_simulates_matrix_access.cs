// vybe-test: csharp/csharp_indexers/two_parameter_indexer_simulates_matrix_access
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

class Matrix{
    double[,] _d=new double[3,3];
    public double this[int r,int c]{get=>_d[r,c]; set=>_d[r,c]=value;}
}
var m=new Matrix(); m[1,2]=9.9;
__P((m[1,2]).ToString());
__Check("9.9");
