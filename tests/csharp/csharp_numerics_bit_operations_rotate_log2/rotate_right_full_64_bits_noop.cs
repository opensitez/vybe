// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/rotate_right_full_64_bits_noop

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

ulong val = 0x123456789ABCDEF0UL;
ulong rot = System.Numerics.BitOperations.RotateRight(val, 64);
__P((rot == val).ToString());
__Check("True");
