// vybe-test: csharp/csharp_iasync_enumerable/async_enumerable_long_stream_count
// origin: languages/csharp/tests/csharp/test_csharp_iasync_enumerable.rs

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

async System.Collections.Generic.IAsyncEnumerable<long> Longs() {
    yield return 100L;
    yield return 200L;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var v in Longs()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("2");
