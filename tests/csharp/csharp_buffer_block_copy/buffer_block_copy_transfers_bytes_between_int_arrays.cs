// vybe-test: csharp/csharp_buffer_block_copy/buffer_block_copy_transfers_bytes_between_int_arrays
// origin: languages/csharp/tests/csharp/test_csharp_buffer_block_copy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] source = { 0x01020304, 0 };
int[] dest = { 0, 0 };
System.Buffer.BlockCopy(source, 0, dest, 0, 4);
__Check((dest[0]).ToString(), "67305985");
