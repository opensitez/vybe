// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_uint32_cases

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

__P(System.Numerics.BitOperations.PopCount(0u).ToString());
__P(System.Numerics.BitOperations.PopCount(1u).ToString());
__P(System.Numerics.BitOperations.PopCount(0b1011u).ToString());
__P(System.Numerics.BitOperations.PopCount(0xFFFFFFFFu).ToString());
__Check("0\n1\n3\n32");
