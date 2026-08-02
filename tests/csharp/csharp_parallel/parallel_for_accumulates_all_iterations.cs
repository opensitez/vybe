// vybe-test: csharp/csharp_parallel/parallel_for_accumulates_all_iterations
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int sum=0;
System.Threading.Tasks.Parallel.For(0,100,i=>{
    System.Threading.Interlocked.Add(ref sum,i);
});
__Check((sum).ToString(), "4950");
