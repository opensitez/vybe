// vybe-test: csharp/csharp_collections_bit_vector_32/bit_vector_32_case_17

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

var bv = new System.Collections.Specialized.BitVector32(17);
int mask = System.Collections.Specialized.BitVector32.CreateMask();
__P(bv.Data.ToString());
__Check("17");
