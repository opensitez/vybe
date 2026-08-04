// vybe-test: csharp/csharp_parallel/parallel_for_each_processes_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_parallel.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var items=new[]{1,2,3,4,5};
int sum=0;
System.Threading.Tasks.Parallel.ForEach(items,n=>{
    System.Threading.Interlocked.Add(ref sum,n);
});
__P((sum).ToString());
__Check("15");
