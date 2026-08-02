// vybe-test: csharp/csharp_linq_aggregate_element/single_with_predicate_zero_matches_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string tag="ok";
try{new[]{1,2,3}.Single(x=>x>10);}catch(System.InvalidOperationException){tag="none";}
__Check((tag).ToString(), "none");
