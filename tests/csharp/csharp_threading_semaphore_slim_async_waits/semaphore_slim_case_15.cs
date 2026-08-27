// vybe-test: csharp/csharp_threading_semaphore_slim_async_waits/semaphore_slim_case_15

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

using var sem = new System.Threading.SemaphoreSlim(1, 1);
bool entered = sem.Wait(100);
__P(entered.ToString());
__P(sem.CurrentCount.ToString());
sem.Release();
__Check("True\n0");
