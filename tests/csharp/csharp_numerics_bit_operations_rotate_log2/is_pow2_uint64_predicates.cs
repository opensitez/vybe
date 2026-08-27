// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/is_pow2_uint64_predicates

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

__P(System.Numerics.BitOperations.IsPow2(0UL).ToString());
__P(System.Numerics.BitOperations.IsPow2(1UL << 50).ToString());
__P(System.Numerics.BitOperations.IsPow2((1UL << 50) + 1).ToString());
__Check("False\nTrue\nFalse");
