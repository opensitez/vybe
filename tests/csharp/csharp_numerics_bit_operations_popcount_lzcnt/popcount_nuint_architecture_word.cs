// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_nuint_architecture_word

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

nuint val = 0b1111;
__P(System.Numerics.BitOperations.PopCount(val).ToString());
__Check("4");
