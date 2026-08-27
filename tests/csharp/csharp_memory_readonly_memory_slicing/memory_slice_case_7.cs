// vybe-test: csharp/csharp_memory_readonly_memory_slicing/memory_slice_case_7

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

ReadOnlyMemory<char> mem = "Hello_World_7".AsMemory();
var slice = mem.Slice(0, 5);
__P(slice.ToString());
__Check("Hello");
