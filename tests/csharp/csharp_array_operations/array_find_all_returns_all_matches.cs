// vybe-test: csharp/csharp_array_operations/array_find_all_returns_all_matches
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] a = {1,2,3,4,5};
int[] evens = System.Array.FindAll(a, x => x%2==0);
__Check((evens.Length).ToString(), "2");
