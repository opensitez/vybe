// vybe-test: csharp/csharp_array_operations/array_copy_transfers_elements_to_destination
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src = {10,20,30}; int[] dst = new int[3];
System.Array.Copy(src, dst, 3);
__Check((dst[1]).ToString(), "20");
