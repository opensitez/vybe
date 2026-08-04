// vybe-test: csharp/csharp_concurrency_sync/reader_writer_lock_allows_multiple_concurrent_readers
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

var rwl=new System.Threading.ReaderWriterLockSlim();
rwl.EnterReadLock();
rwl.EnterReadLock();
__P((rwl.CurrentReadCount).ToString());
rwl.ExitReadLock();
rwl.ExitReadLock();
__Check("2");
