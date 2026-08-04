// vybe-test: csharp/csharp_concurrency_sync/lock_statement_serialises_access_to_shared_counter
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

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

int counter=0;
object lk=new object();
var tasks=new System.Threading.Tasks.Task[10];
for(int i=0;i<10;i++){
    tasks[i]=System.Threading.Tasks.Task.Run(()=>{lock(lk){counter++;}});
}
System.Threading.Tasks.Task.WaitAll(tasks);
__P((counter).ToString());
__Check("10");
