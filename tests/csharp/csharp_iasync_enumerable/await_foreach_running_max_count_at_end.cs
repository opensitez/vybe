// vybe-test: csharp/csharp_iasync_enumerable/await_foreach_running_max_count_at_end
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

async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 3;
    yield return 7;
    yield return 5;
    yield return 9;
}
async System.Threading.Tasks.Task Run() {
    int max = int.MinValue;
    int count = 0;
    await foreach (var x in Stream()) { if (x > max) max = x; count++; }
    __P((count).ToString());
}
Run().Wait();
__Check("4");
