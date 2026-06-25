//! Synchronization primitives: `lock`, `Monitor`, `SemaphoreSlim`, `ReaderWriterLockSlim`.
use super::helpers::run_csharp;

#[test]
fn lock_statement_serialises_access_to_shared_counter() {
    assert_eq!(
        run_csharp(r#"int counter=0;
object lk=new object();
var tasks=new System.Threading.Tasks.Task[10];
for(int i=0;i<10;i++){
    tasks[i]=System.Threading.Tasks.Task.Run(()=>{lock(lk){counter++;}});
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);"#),
        &["10"]
    );
}

#[test]
fn semaphore_slim_limits_concurrent_entries() {
    assert_eq!(
        run_csharp(r#"var sem=new System.Threading.SemaphoreSlim(1,1);
sem.Wait();
Console.WriteLine(sem.CurrentCount);
sem.Release();
Console.WriteLine(sem.CurrentCount);"#),
        &["0", "1"]
    );
}

#[test]
fn monitor_try_enter_returns_false_when_already_locked() {
    assert_eq!(
        run_csharp(r#"object obj=new object();
System.Threading.Monitor.Enter(obj);
bool got=System.Threading.Monitor.TryEnter(obj,0);
System.Threading.Monitor.Exit(obj);
Console.WriteLine(got);"#),
        &["False"]
    );
}

#[test]
fn reader_writer_lock_allows_multiple_concurrent_readers() {
    assert_eq!(
        run_csharp(r#"var rwl=new System.Threading.ReaderWriterLockSlim();
rwl.EnterReadLock();
rwl.EnterReadLock();
Console.WriteLine(rwl.CurrentReadCount);
rwl.ExitReadLock();
rwl.ExitReadLock();"#),
        &["2"]
    );
}

#[test]
fn interlocked_compare_exchange_sets_only_when_expected() {
    assert_eq!(
        run_csharp(r#"int val=0;
int original=System.Threading.Interlocked.CompareExchange(ref val,99,0);
Console.WriteLine(original); Console.WriteLine(val);"#),
        &["0", "99"]
    );
}
