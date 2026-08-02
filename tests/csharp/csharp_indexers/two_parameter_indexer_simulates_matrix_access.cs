// vybe-test: csharp/csharp_indexers/two_parameter_indexer_simulates_matrix_access
// origin: languages/csharp/tests/csharp/test_csharp_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Matrix{
    double[,] _d=new double[3,3];
    public double this[int r,int c]{get=>_d[r,c]; set=>_d[r,c]=value;}
}
var m=new Matrix(); m[1,2]=9.9;
__Check((m[1,2]).ToString(), "9.9");
