// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/leading_zero_count_hex_boundaries

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

__P(System.Numerics.BitOperations.LeadingZeroCount(0x0FFFFFFFu).ToString());
__Check("4");
