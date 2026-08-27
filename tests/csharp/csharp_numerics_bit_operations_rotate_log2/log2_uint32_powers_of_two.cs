// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/log2_uint32_powers_of_two

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

__P(System.Numerics.BitOperations.Log2(1u).ToString());
__P(System.Numerics.BitOperations.Log2(2u).ToString());
__P(System.Numerics.BitOperations.Log2(64u).ToString());
__P(System.Numerics.BitOperations.Log2(1024u).ToString());
__Check("0\n1\n6\n10");
