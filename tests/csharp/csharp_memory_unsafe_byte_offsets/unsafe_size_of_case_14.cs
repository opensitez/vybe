// vybe-test: csharp/csharp_memory_unsafe_byte_offsets/unsafe_size_of_case_14

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

int size = System.Runtime.CompilerServices.Unsafe.SizeOf<int>();
__P(size.ToString());
__Check("4");
