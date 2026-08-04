// vybe-test: csharp/csharp_async_advanced/value_task_from_result_avoids_allocation
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

async System.Threading.Tasks.ValueTask<int> GetValueAsync()=>42;
int v=await GetValueAsync();
__P((v).ToString());
__Check("42");
