// vybe-test: csharp/csharp_task_combinators/when_all_seven_incremental_sum
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        N(1), N(2), N(3), N(4), N(5), N(6), N(7)
    );
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
