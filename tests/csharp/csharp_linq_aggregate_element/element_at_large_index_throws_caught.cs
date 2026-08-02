// vybe-test: csharp/csharp_linq_aggregate_element/element_at_large_index_throws_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string tag="ok";
try{new[]{1,2}.ElementAt(5);}catch(System.ArgumentOutOfRangeException){tag="range";}
__Check((tag).ToString(), "range");
