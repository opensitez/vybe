// vybe-test: csharp/csharp_linq_aggregate_element/single_empty_throws_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string tag="ok";
try{System.Array.Empty<int>().Single();}catch(System.InvalidOperationException){tag="empty";}
__Check((tag).ToString(), "empty");
