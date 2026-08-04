// vybe-test: csharp/csharp_iasync_enumerable/sequential_await_foreach_on_fresh_factory_counts_twice
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

async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int total = 0;
    await foreach (var x in Make()) total++;
    await foreach (var x in Make()) total++;
    __P((total).ToString());
}
Run().Wait();
__Check("4");
