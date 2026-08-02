// vybe-test: csharp/linq_lambdas/linq_firstordefault_empty
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums = new List<int>();
__Check((nums.FirstOrDefault()).ToString(), "0");
