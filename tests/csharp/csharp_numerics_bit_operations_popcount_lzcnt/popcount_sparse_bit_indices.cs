// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_sparse_bit_indices

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

uint sparse = (1u << 3) | (1u << 15) | (1u << 29);
__P(System.Numerics.BitOperations.PopCount(sparse).ToString());
__Check("3");
