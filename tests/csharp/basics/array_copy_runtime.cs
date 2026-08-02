// vybe-test: csharp/basics/array_copy_runtime
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src = new int[] { 10, 20, 30, 40 };
int[] dst = new int[] { 0, 0, 0, 0 };
Array.Copy(src, dst, 3);
__Check((dst[0] + dst[1] + dst[2] + dst[3]).ToString(), "60");
