// vybe-test: csharp/csharp_memory_binary_primitives_endianness/binary_primitives_endian_case_6

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

byte[] bytes = new byte[4];
System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(bytes, 6);
int val = System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(bytes);
__P(val.ToString());
__Check("6");
