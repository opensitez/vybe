// vybe-test: csharp/csharp_async_await_flow/task_run_offloads_work_and_returns_result_to_awaiter
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
async Task<int> Run() {
    return await Task.Run(() => 11);
}
__P((Run().GetAwaiter().GetResult()).ToString());
__Check("11");
