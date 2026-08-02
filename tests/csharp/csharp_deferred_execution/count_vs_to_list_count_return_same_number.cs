// vybe-test: csharp/csharp_deferred_execution/count_vs_to_list_count_return_same_number
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var q=new[]{1,2,3,4}.Where(x=>x%2==0);
__Check((q.Count()).ToString(), "2");
__Check((q.ToList().Count).ToString(), "2");
