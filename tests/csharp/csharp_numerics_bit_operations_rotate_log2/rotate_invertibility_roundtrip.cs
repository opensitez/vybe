// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/rotate_invertibility_roundtrip

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

uint orig = 0xCAFEBABE;
uint rot = System.Numerics.BitOperations.RotateLeft(orig, 13);
uint recovered = System.Numerics.BitOperations.RotateRight(rot, 13);
__P((recovered == orig).ToString());
__Check("True");
