// vybe-test: csharp/csharp_async_await_flow/await_in_try_finally_still_runs_finally_before_result_is_observed
// origin: languages/csharp/tests/csharp/test_csharp_async_await_flow.rs

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

using System.Threading.Tasks;
async Task<int> Pick() {
    try {
        return await Task.FromResult(2);
    } finally {
        __P(("cleanup").ToString());
    }
}
__P((Pick().GetAwaiter().GetResult()).ToString());
__Check("cleanup\n2");
