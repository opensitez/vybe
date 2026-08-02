// vybe-test: csharp/csharp_task_combinators/when_all_then_continue_with_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.WhenAll(N(2), N(3))
        .ContinueWith(t => {
            foreach (var x in t.Result) count += x;
        });
    Console.WriteLine(count);
}
Run().Wait();
