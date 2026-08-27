// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/is_pow2_uint32_predicates

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

__P(System.Numerics.BitOperations.IsPow2(0u).ToString());
__P(System.Numerics.BitOperations.IsPow2(1u).ToString());
__P(System.Numerics.BitOperations.IsPow2(16u).ToString());
__P(System.Numerics.BitOperations.IsPow2(18u).ToString());
__Check("False\nTrue\nTrue\nFalse");
