// vybe-test: csharp/csharp_concurrency_sync/reader_writer_lock_allows_multiple_concurrent_readers
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var rwl=new System.Threading.ReaderWriterLockSlim();
rwl.EnterReadLock();
rwl.EnterReadLock();
__Check((rwl.CurrentReadCount).ToString(), "2");
rwl.ExitReadLock();
rwl.ExitReadLock();
