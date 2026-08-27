// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/trailing_zero_count_uint64

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

__P(System.Numerics.BitOperations.TrailingZeroCount(0UL).ToString());
__P(System.Numerics.BitOperations.TrailingZeroCount(16UL).ToString());
__Check("64\n4");
