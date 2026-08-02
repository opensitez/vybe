// vybe-test: csharp/csharp_deferred_execution/any_short_circuits_after_first_match
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count=0;
bool found=new[]{1,2,3,4,5}.Any(n=>{count++;return n==3;});
__Check((found).ToString(), "True"); __Check((count).ToString(), "3");
