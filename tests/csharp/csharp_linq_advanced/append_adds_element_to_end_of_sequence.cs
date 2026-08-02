// vybe-test: csharp/csharp_linq_advanced/append_adds_element_to_end_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{1,2,3}.Append(4);
__Check((result.Last()).ToString(), "4");
