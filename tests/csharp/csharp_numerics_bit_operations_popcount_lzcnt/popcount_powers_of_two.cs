// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_powers_of_two

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

for (int i = 0; i < 5; i++) {
    uint p = 1u << i;
    __P(System.Numerics.BitOperations.PopCount(p).ToString());
}
__Check("1\n1\n1\n1\n1");
