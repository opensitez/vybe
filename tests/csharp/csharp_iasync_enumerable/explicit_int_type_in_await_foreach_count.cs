// vybe-test: csharp/csharp_iasync_enumerable/explicit_int_type_in_await_foreach_count
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
    yield return 4;
    yield return 8;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (int x in Stream()) count++;
    __P((count).ToString());
}
Run().Wait();
__Check("2");
