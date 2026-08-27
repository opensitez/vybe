// vybe-test: csharp/csharp_memory_array_pool_rent_return/array_pool_case_18

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

byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(180);
__P((buf.Length >= 180).ToString());
System.Buffers.ArrayPool<byte>.Shared.Return(buf);
__Check("True");
