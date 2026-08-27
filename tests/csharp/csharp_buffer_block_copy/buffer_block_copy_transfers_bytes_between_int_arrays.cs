// vybe-test: csharp/csharp_buffer_block_copy/buffer_block_copy_transfers_bytes_between_int_arrays
// origin: languages/csharp/tests/csharp/test_csharp_buffer_block_copy.rs

using static __Harness;

int[] source = { 0x01020304, 0 }
;
int[] dest = { 0, 0 }
;
System.Buffer.BlockCopy(source, 0, dest, 0, 4);
__P((dest[0]).ToString());
__Check("16909060");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
