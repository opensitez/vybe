//! Task combinators — `Task.WhenAll`, `Task.WhenAny`, `ContinueWith`, `Task.Run`.
//! GAP: concurrency primitives.

csharp_cases! {
    task_run_returns_constant_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => 9);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["9"]
    };

    task_run_computes_sum_before_return => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => 3 + 4);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["7"]
    };

    task_run_captures_local_and_doubles => {
        r#"
async System.Threading.Tasks.Task Run() {
    int seed = 5;
    var t = System.Threading.Tasks.Task.Run(() => seed * 2);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["10"]
    };

    task_run_string_length_as_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => "hello".Length);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["5"]
    };

    task_run_nested_addition_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var outer = System.Threading.Tasks.Task.Run(() => {
        return System.Threading.Tasks.Task.Run(() => 2 + 3).Result;
    });
    Console.WriteLine(outer.Result);
}
Run().Wait();
"#,
        ["5"]
    };

    task_run_void_increments_counter => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    System.Threading.Tasks.Task.Run(() => { count = 4; }).Wait();
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    task_run_three_parallel_workers_sum => {
        r#"
async System.Threading.Tasks.Task Run() {
    var a = System.Threading.Tasks.Task.Run(() => 1);
    var b = System.Threading.Tasks.Task.Run(() => 2);
    var c = System.Threading.Tasks.Task.Run(() => 3);
    Console.WriteLine(a.Result + b.Result + c.Result);
}
Run().Wait();
"#,
        ["6"]
    };

    when_all_two_tasks_sum_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(4), N(6));
    Console.WriteLine(results[0] + results[1]);
}
Run().Wait();
"#,
        ["10"]
    };

    when_all_three_tasks_sum_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(1), N(2), N(3));
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    when_all_four_tasks_product_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(2), N(3), N(4), N(5));
    int count = 1;
    foreach (var x in results) count *= x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["120"]
    };

    when_all_five_tasks_length_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(1), N(2), N(3), N(4), N(5));
    Console.WriteLine(results.Length);
}
Run().Wait();
"#,
        ["5"]
    };

    when_all_with_yield_preserves_order_sum => {
        r#"
async System.Threading.Tasks.Task<int> Val(int n) {
    await System.Threading.Tasks.Task.Yield();
    return n;
}
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(Val(10), Val(20));
    Console.WriteLine(results[0] + results[1]);
}
Run().Wait();
"#,
        ["30"]
    };

    when_all_from_result_tasks_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.FromResult(7),
        System.Threading.Tasks.Task.FromResult(8)
    );
    Console.WriteLine(results[0] + results[1]);
}
Run().Wait();
"#,
        ["15"]
    };

    when_all_task_run_mixed_sum => {
        r#"
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.Run(() => 2),
        System.Threading.Tasks.Task.Run(() => 5)
    );
    Console.WriteLine(results[0] + results[1]);
}
Run().Wait();
"#,
        ["7"]
    };

    when_all_empty_array_length_zero => {
        r#"
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(new System.Threading.Tasks.Task<int>[0]);
    Console.WriteLine(results.Length);
}
Run().Wait();
"#,
        ["0"]
    };

    when_all_single_task_returns_singleton => {
        r#"
async System.Threading.Tasks.Task<int> Solo() { return 11; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(Solo());
    Console.WriteLine(results[0]);
}
Run().Wait();
"#,
        ["11"]
    };

    when_all_six_tasks_sum_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(1), N(1), N(1), N(1), N(1), N(1));
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    when_all_negative_values_sum => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(-2), N(5), N(-1));
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    when_all_zero_values_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(0), N(0), N(0));
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["0"]
    };

    when_any_picks_from_result_fast_task => {
        r#"
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.FromResult(3),
        System.Threading.Tasks.Task.FromResult(9)
    );
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["3"]
    };

    when_any_two_identical_fast_tasks_count => {
        r#"
async System.Threading.Tasks.Task<int> Fast() { return 4; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Fast());
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["4"]
    };

    when_any_fast_beats_delayed_task => {
        r#"
async System.Threading.Tasks.Task<int> Fast() { return 1; }
async System.Threading.Tasks.Task<int> Slow() {
    await System.Threading.Tasks.Task.Delay(1000);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Fast(), Slow());
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["1"]
    };

    when_any_three_tasks_first_completed_count => {
        r#"
async System.Threading.Tasks.Task<int> A() { return 10; }
async System.Threading.Tasks.Task<int> B() { return 20; }
async System.Threading.Tasks.Task<int> C() { return 30; }
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B(), C());
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["10"]
    };

    when_any_with_task_run_winner_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.Run(() => 6),
        System.Threading.Tasks.Task.Run(() => 7)
    );
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["6"]
    };

    when_any_yield_then_return_count => {
        r#"
async System.Threading.Tasks.Task<int> Yielded() {
    await System.Threading.Tasks.Task.Yield();
    return 8;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Yielded(), Yielded());
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["8"]
    };

    when_any_loser_not_awaited_count_stays_one => {
        r#"
async System.Threading.Tasks.Task<int> Win() { return 1; }
async System.Threading.Tasks.Task<int> Lose() {
    await System.Threading.Tasks.Task.Delay(500);
    return 99;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Win(), Lose());
    int count = winner.Result;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    continue_with_doubles_task_result => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => 7)
        .ContinueWith(t => count = t.Result * 2);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["14"]
    };

    continue_with_adds_one_to_result => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(5)
        .ContinueWith(t => count = t.Result + 1);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    continue_with_chained_two_steps => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(2)
        .ContinueWith(t => t.Result + 3)
        .ContinueWith(t => count = t.Result * 4);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["20"]
    };

    continue_with_on_task_run_preserves_value => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => 12)
        .ContinueWith(t => count = t.Result);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["12"]
    };

    continue_with_void_task_sets_flag_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => { })
        .ContinueWith(t => count = 1);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    continue_with_three_link_chain_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(1)
        .ContinueWith(t => t.Result + 1)
        .ContinueWith(t => t.Result + 1)
        .ContinueWith(t => count = t.Result + 1);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    continue_with_squares_result => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(5)
        .ContinueWith(t => count = t.Result * t.Result);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["25"]
    };

    when_all_then_continue_with_sum => {
        r#"
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
"#,
        ["5"]
    };

    task_run_then_continue_with_increment => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.Run(() => 9)
        .ContinueWith(t => count = t.Result + 1);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["10"]
    };

    when_all_eight_ones_count => {
        r#"
async System.Threading.Tasks.Task<int> One() { return 1; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        One(), One(), One(), One(), One(), One(), One(), One()
    );
    Console.WriteLine(results.Length);
}
Run().Wait();
"#,
        ["8"]
    };

    when_all_task_run_four_workers_sum => {
        r#"
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(
        System.Threading.Tasks.Task.Run(() => 1),
        System.Threading.Tasks.Task.Run(() => 2),
        System.Threading.Tasks.Task.Run(() => 3),
        System.Threading.Tasks.Task.Run(() => 4)
    );
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["10"]
    };

    when_any_four_delayed_vs_immediate_count => {
        r#"
async System.Threading.Tasks.Task<int> Now() { return 2; }
async System.Threading.Tasks.Task<int> Later() {
    await System.Threading.Tasks.Task.Delay(200);
    return 3;
}
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(Now(), Later(), Later(), Later());
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["2"]
    };

    when_all_preserves_first_element => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(42), N(99));
    Console.WriteLine(results[0]);
}
Run().Wait();
"#,
        ["42"]
    };

    when_all_preserves_last_element => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(42), N(99));
    Console.WriteLine(results[1]);
}
Run().Wait();
"#,
        ["99"]
    };

    task_run_bool_true_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => true);
    Console.WriteLine(t.Result ? 1 : 0);
}
Run().Wait();
"#,
        ["1"]
    };

    task_run_bool_false_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => false);
    Console.WriteLine(t.Result ? 1 : 0);
}
Run().Wait();
"#,
        ["0"]
    };

    when_all_max_of_three_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(3), N(9), N(5));
    int max = results[0];
    foreach (var x in results) if (x > max) max = x;
    Console.WriteLine(max);
}
Run().Wait();
"#,
        ["9"]
    };

    when_all_min_of_three_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(3), N(9), N(5));
    int min = results[0];
    foreach (var x in results) if (x < min) min = x;
    Console.WriteLine(min);
}
Run().Wait();
"#,
        ["3"]
    };

    continue_with_from_when_any_result => {
        r#"
async System.Threading.Tasks.Task<int> A() { return 4; }
async System.Threading.Tasks.Task<int> B() { return 8; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    var winner = await System.Threading.Tasks.Task.WhenAny(A(), B());
    await winner.ContinueWith(t => count = t.Result + 1);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    task_run_modulo_result_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => 17 % 5);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["2"]
    };

    when_all_with_delay_zero_sum => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) {
    await System.Threading.Tasks.Task.Delay(0);
    return v;
}
async System.Threading.Tasks.Task Run() {
    var results = await System.Threading.Tasks.Task.WhenAll(N(2), N(3));
    Console.WriteLine(results[0] + results[1]);
}
Run().Wait();
"#,
        ["5"]
    };

    continue_with_uses_previous_task_status_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(1)
        .ContinueWith(t => count = t.IsCompleted ? 1 : 0);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    when_any_task_run_vs_from_result => {
        r#"
async System.Threading.Tasks.Task Run() {
    var winner = await System.Threading.Tasks.Task.WhenAny(
        System.Threading.Tasks.Task.Run(() => 15),
        System.Threading.Tasks.Task.FromResult(16)
    );
    Console.WriteLine(winner.Result);
}
Run().Wait();
"#,
        ["15"]
    };

    task_run_array_length_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => new int[] { 1, 2, 3, 4 }.Length);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["4"]
    };

    when_all_loop_spawned_tasks_count => {
        r#"
async System.Threading.Tasks.Task<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    var tasks = new System.Threading.Tasks.Task<int>[3];
    for (int i = 0; i < 3; i++) tasks[i] = N(i + 1);
    var results = await System.Threading.Tasks.Task.WhenAll(tasks);
    int count = 0;
    foreach (var x in results) count += x;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    continue_with_sequential_from_result_sum => {
        r#"
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await System.Threading.Tasks.Task.FromResult(3)
        .ContinueWith(t => count += t.Result);
    await System.Threading.Tasks.Task.FromResult(4)
        .ContinueWith(t => count += t.Result);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["7"]
    };

    when_all_seven_incremental_sum => {
        r#"
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
"#,
        ["28"]
    };

    task_run_negative_result_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var t = System.Threading.Tasks.Task.Run(() => -12);
    Console.WriteLine(t.Result);
}
Run().Wait();
"#,
        ["-12"]
    };

    when_any_continue_with_doubles_winner => {
        r#"
async System.Threading.Tasks.Task<int> Win() { return 6; }
async System.Threading.Tasks.Task<int> Lose() {
    await System.Threading.Tasks.Task.Delay(300);
    return 1;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    var winner = await System.Threading.Tasks.Task.WhenAny(Win(), Lose());
    await winner.ContinueWith(t => count = t.Result * 2);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["12"]
    };
}
