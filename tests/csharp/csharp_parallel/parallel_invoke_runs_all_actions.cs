// vybe-test: csharp/csharp_parallel/parallel_invoke_runs_all_actions
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=0,b=0;
System.Threading.Tasks.Parallel.Invoke(
    ()=>System.Threading.Interlocked.Exchange(ref a,1),
    ()=>System.Threading.Interlocked.Exchange(ref b,2)
);
__Check((a+b).ToString(), "3");
