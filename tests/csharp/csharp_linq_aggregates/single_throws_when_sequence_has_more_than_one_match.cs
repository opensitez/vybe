// vybe-test: csharp/csharp_linq_aggregates/single_throws_when_sequence_has_more_than_one_match
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "ok";
try { new[]{1,2}.Single(); }
catch(System.InvalidOperationException) { result = "many"; }
__Check((result).ToString(), "many");
