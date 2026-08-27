// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_loop_over_ints

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

int sum = 0;
for (uint i = 1; i <= 3; i++) {
    sum += System.Numerics.BitOperations.PopCount(i); // 1 + 1 + 2 = 4
}
__P(sum.ToString());
__Check("4");
