// vybe-test: csharp/csharp_task_combinators/when_all_min_of_three_count
// origin: languages/csharp/tests/csharp/test_csharp_task_combinators.rs

async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(3), N(9), N(5));
    int min = results[0];
    foreach (var x in results) if (x < min) min = x;
    Console.WriteLine(min);
}
Run().Wait();
