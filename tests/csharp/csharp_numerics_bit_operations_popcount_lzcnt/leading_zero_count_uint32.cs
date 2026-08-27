// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/leading_zero_count_uint32

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

__P(System.Numerics.BitOperations.LeadingZeroCount(0u).ToString());
__P(System.Numerics.BitOperations.LeadingZeroCount(1u).ToString());
__P(System.Numerics.BitOperations.LeadingZeroCount(0x80000000u).ToString());
__Check("32\n31\n0");
