// vybe-test: csharp/csharp_memory_pinned_gc_handles/pinned_gc_handle_case_12

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

byte[] arr = new byte[] { (byte)12, 2, 3 };
var handle = System.Runtime.InteropServices.GCHandle.Alloc(arr, System.Runtime.InteropServices.GCHandleType.Pinned);
__P(handle.IsAllocated.ToString());
handle.Free();
__Check("True");
