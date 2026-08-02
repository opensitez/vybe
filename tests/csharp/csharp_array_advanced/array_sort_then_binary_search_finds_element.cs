// vybe-test: csharp/csharp_array_advanced/array_sort_then_binary_search_finds_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={5,3,1,4,2};
System.Array.Sort(arr);
int idx=System.Array.BinarySearch(arr,4);
__Check((idx).ToString(), "3");
