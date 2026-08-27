// vybe-test: csharp/csharp_memory_marshal_casts/memory_marshal_case_16

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

byte[] bytes = new byte[] { (byte)16, 0, 0, 0 };
ReadOnlySpan<int> ints = System.Runtime.InteropServices.MemoryMarshal.Cast<byte, int>(bytes);
__P(ints.Length.ToString());
__P(ints[0].ToString());
__Check("1\n16");
