// vybe-test: csharp/csharp_linq_advanced/prepend_adds_element_to_start_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{2,3,4}.Prepend(1);
__Check((result.First()).ToString(), "1");
