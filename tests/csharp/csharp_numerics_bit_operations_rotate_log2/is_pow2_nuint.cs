// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/is_pow2_nuint

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

nuint v1 = 512;
nuint v2 = 500;
__P(System.Numerics.BitOperations.IsPow2(v1).ToString());
__P(System.Numerics.BitOperations.IsPow2(v2).ToString());
__Check("True\nFalse");
