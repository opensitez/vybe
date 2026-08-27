// vybe-test: csharp/csharp_memory_native_memory_alloc_free/native_memory_case_4

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

IntPtr ptr = System.Runtime.InteropServices.Marshal.AllocHGlobal(32);
__P((ptr != IntPtr.Zero).ToString());
System.Runtime.InteropServices.Marshal.FreeHGlobal(ptr);
__Check("True");
