// vybe-test: csharp/csharp_threading_reader_writer_lock_slim/reader_writer_lock_case_19

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var rw = new System.Threading.ReaderWriterLockSlim();
rw.EnterReadLock();
__P(rw.IsReadLockHeld.ToString());
rw.ExitReadLock();
__Check("True");
