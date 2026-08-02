// vybe-test: csharp/csharp_deferred_execution/where_query_not_executed_until_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count=0;
var q=new[]{1,2,3}.Where(n=>{count++;return n>1;});
__Check((count).ToString(), "0");
var list=q.ToList();
__Check((count).ToString(), "3");
