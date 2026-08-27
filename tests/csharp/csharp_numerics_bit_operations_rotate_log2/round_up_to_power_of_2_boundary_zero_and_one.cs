// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/round_up_to_power_of_2_boundary_zero_and_one

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

__P(System.Numerics.BitOperations.RoundUpToPowerOf2(0u).ToString());
__P(System.Numerics.BitOperations.RoundUpToPowerOf2(1u).ToString());
__Check("0\n1");
