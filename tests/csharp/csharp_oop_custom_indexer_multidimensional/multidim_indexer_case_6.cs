// vybe-test: csharp/csharp_oop_custom_indexer_multidimensional/multidim_indexer_case_6

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

var m = new MatrixStore_6();
m[0, 1] = 6;
__P(m[0, 1].ToString());
__Check("6");

class MatrixStore_6 {
    private int[,] data = new int[2, 2];
    public int this[int r, int c] {
        get => data[r, c];
        set => data[r, c] = value;
    }
}
