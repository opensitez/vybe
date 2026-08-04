// vybe-test: csharp/csharp_parallel/parallel_for_accumulates_all_iterations
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

int sum=0;
System.Threading.Tasks.Parallel.For(0,100,i=>{
    System.Threading.Interlocked.Add(ref sum,i);
});
__P((sum).ToString());
__Check("4950");
