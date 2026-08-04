// vybe-test: csharp/csharp_buffer_block_copy/buffer_block_copy_transfers_bytes_between_int_arrays
// origin: languages/csharp/tests/csharp/test_csharp_buffer_block_copy.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int[] source = { 0x01020304, 0 };
int[] dest = { 0, 0 };
System.Buffer.BlockCopy(source, 0, dest, 0, 4);
__P((dest[0]).ToString());
__Check("67305985");
