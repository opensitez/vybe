// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_uint64_alternating_bits

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

ulong alt = 0xAAAAAAAAAAAAAAAAUL; // 32 set bits
__P(System.Numerics.BitOperations.PopCount(alt).ToString());
__Check("32");
