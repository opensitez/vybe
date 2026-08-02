// vybe-test: csharp/csharp_concurrency_sync/lock_statement_serialises_access_to_shared_counter
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

int counter=0;
object lk=new object();
var tasks=new System.Threading.Tasks.Task[10];
for(int i=0;i<10;i++){
    tasks[i]=System.Threading.Tasks.Task.Run(()=>{lock(lk){counter++;}});
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
