// vybe-test: csharp/csharp_numerics_bit_operations_rotate_log2/log2_non_powers_of_two_floor

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

__P(System.Numerics.BitOperations.Log2(7u).ToString());
__P(System.Numerics.BitOperations.Log2(9u).ToString());
__Check("2\n3");
