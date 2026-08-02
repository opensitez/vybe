// vybe-test: csharp/csharp_array_advanced/array_convert_all_transforms_each_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src={1,2,3};
string[] dst=System.Array.ConvertAll(src,n=>n.ToString()+"x");
__Check((dst[1]).ToString(), "2x");
