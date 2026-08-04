// vybe-test: csharp/csharp_async_advanced/when_all_awaits_all_tasks_and_returns_results
// origin: languages/csharp/tests/csharp/test_csharp_async_advanced.rs

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

async System.Threading.Tasks.Task<int> N(int v){
    await System.Threading.Tasks.Task.Delay(0);return v;
}
int[] results=await System.Threading.Tasks.Task.WhenAll(N(1),N(2),N(3));
__P((results.Sum()).ToString());
__Check("6");
