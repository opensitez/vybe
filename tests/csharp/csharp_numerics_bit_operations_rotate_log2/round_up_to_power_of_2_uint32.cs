// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/round_up_to_power_of_2_uint32

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

__P(System.Numerics.BitOperations.RoundUpToPowerOf2(5u).ToString());
__P(System.Numerics.BitOperations.RoundUpToPowerOf2(16u).ToString());
__P(System.Numerics.BitOperations.RoundUpToPowerOf2(17u).ToString());
__Check("8\n16\n32");
