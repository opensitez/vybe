// vybe-test: csharp/csharp_array_advanced/array_fill_sets_all_elements_to_value
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr=new int[5];
System.Array.Fill(arr,7);
__Check((arr[2]).ToString(), "7");
