// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/leading_trailing_sum_zero_input

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

uint z = 0u;
__P(System.Numerics.BitOperations.LeadingZeroCount(z).ToString());
__P(System.Numerics.BitOperations.TrailingZeroCount(z).ToString());
__Check("32\n32");
