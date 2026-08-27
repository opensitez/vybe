// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/rotate_left_uint64

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

ulong val = 0x8000000000000001UL;
ulong rot = System.Numerics.BitOperations.RotateLeft(val, 1);
__P(rot.ToString("X"));
__Check("3");
