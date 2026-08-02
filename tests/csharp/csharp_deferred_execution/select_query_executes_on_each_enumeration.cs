// vybe-test: csharp/csharp_deferred_execution/select_query_executes_on_each_enumeration
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int calls=0;
var q=new[]{1,2,3}.Select(n=>{calls++;return n*2;});
var r1=q.ToList(); var r2=q.ToList();
__Check((calls).ToString(), "6");
