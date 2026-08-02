// vybe-test: csharp/csharp_parallel/parallel_for_each_processes_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var items=new[]{1,2,3,4,5};
int sum=0;
System.Threading.Tasks.Parallel.ForEach(items,n=>{
    System.Threading.Interlocked.Add(ref sum,n);
});
__Check((sum).ToString(), "15");
