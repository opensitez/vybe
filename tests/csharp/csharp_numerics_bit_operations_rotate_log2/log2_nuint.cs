// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/log2_nuint

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

nuint val = 256;
__P(System.Numerics.BitOperations.Log2(val).ToString());
__Check("8");
