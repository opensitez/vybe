// vybe-test: csharp/csharp_numerics_bit_operations_popcount_lzcnt/popcount_single_byte_values

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

uint b1 = 0x0F; // 4
uint b2 = 0xFF; // 8
__P(System.Numerics.BitOperations.PopCount(b1).ToString());
__P(System.Numerics.BitOperations.PopCount(b2).ToString());
__Check("4\n8");
