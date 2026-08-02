// vybe-test: csharp/csharp_linq_chaining/zip_three_sequences_pairwise_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new[]{1,2,3}; var b=new[]{10,20,30};
var result=a.Zip(b).Select(t=>t.First+t.Second);
__Check((string.Join(",",result)).ToString(), "11,22,33");
