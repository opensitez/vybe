// vybe-test: csharp/csharp_async_advanced/async_stream_yields_values_to_await_foreach
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

async System.Collections.Generic.IAsyncEnumerable<int> Sequence(){
    for(int i=1;i<=3;i++){
        await System.Threading.Tasks.Task.Yield();
        yield return i;
    }
}
int sum=0;
await foreach(var n in Sequence()) sum+=n;
__P((sum).ToString());
__Check("6");
