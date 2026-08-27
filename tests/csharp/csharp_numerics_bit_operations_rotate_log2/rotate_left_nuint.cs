// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/rotate_left_nuint

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

nuint val = 1;
nuint rot = System.Numerics.BitOperations.RotateLeft(val, 4);
__P(rot.ToString());
__Check("16");
