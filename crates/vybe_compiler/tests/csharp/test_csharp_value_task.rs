//! `ValueTask` and `ValueTask<T>` — lightweight async completion and `.AsTask()`.

csharp_cases! {
    value_task_int_synchronous_completion => {
        r#"
async System.Threading.Tasks.ValueTask<int> Get() { return 42; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Get());
}
Run().Wait();
"#,
        ["42"]
    };

    value_task_bool_true_result => {
        r#"
async System.Threading.Tasks.ValueTask<bool> Flag() { return true; }
async System.Threading.Tasks.Task Run() {
    bool v = await Flag();
    Console.WriteLine(v ? 1 : 0);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_string_length_as_count => {
        r#"
async System.Threading.Tasks.ValueTask<string> Name() { return "hello"; }
async System.Threading.Tasks.Task Run() {
    string s = await Name();
    Console.WriteLine(s.Length);
}
Run().Wait();
"#,
        ["5"]
    };

    value_task_from_result_returns_value => {
        r#"
async System.Threading.Tasks.Task Run() {
    var vt = System.Threading.Tasks.ValueTask.FromResult(17);
    Console.WriteLine(await vt);
}
Run().Wait();
"#,
        ["17"]
    };

    value_task_completed_task_await_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    await System.Threading.Tasks.ValueTask.CompletedTask;
    Console.WriteLine(1);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_as_task_preserves_int_result => {
        r#"
async System.Threading.Tasks.ValueTask<int> Get() { return 88; }
async System.Threading.Tasks.Task Run() {
    var task = Get().AsTask();
    Console.WriteLine(await task);
}
Run().Wait();
"#,
        ["88"]
    };

    async_value_task_without_yield_sync_path => {
        r#"
async System.Threading.Tasks.ValueTask<int> Compute() { return 3 + 4; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Compute());
}
Run().Wait();
"#,
        ["7"]
    };

    async_value_task_with_yield_then_return => {
        r#"
async System.Threading.Tasks.ValueTask<int> Compute() {
    await System.Threading.Tasks.Task.Yield();
    return 9;
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Compute());
}
Run().Wait();
"#,
        ["9"]
    };

    nested_value_task_await_chain => {
        r#"
async System.Threading.Tasks.ValueTask<int> Inner() { return 5; }
async System.Threading.Tasks.ValueTask<int> Outer() { return await Inner() + 1; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Outer());
}
Run().Wait();
"#,
        ["6"]
    };

    value_task_int_arithmetic_after_await => {
        r#"
async System.Threading.Tasks.ValueTask<int> Base() { return 10; }
async System.Threading.Tasks.Task Run() {
    int v = await Base();
    Console.WriteLine(v * 3);
}
Run().Wait();
"#,
        ["30"]
    };

    sequential_value_task_await_sum_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> N(int x) { return x; }
async System.Threading.Tasks.Task Run() {
    int total = await N(1) + await N(2) + await N(3);
    Console.WriteLine(total);
}
Run().Wait();
"#,
        ["6"]
    };

    value_task_loop_increment_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    for (int i = 0; i < 5; i++) count += await Step();
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    value_task_long_result => {
        r#"
async System.Threading.Tasks.ValueTask<long> Get() { return 1000000L; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Get());
}
Run().Wait();
"#,
        ["1000000"]
    };

    value_task_double_result => {
        r#"
async System.Threading.Tasks.ValueTask<double> Get() { return 3.5; }
async System.Threading.Tasks.Task Run() {
    double v = await Get();
    Console.WriteLine((int)v);
}
Run().Wait();
"#,
        ["3"]
    };

    value_task_zero_result => {
        r#"
async System.Threading.Tasks.ValueTask<int> Zero() { return 0; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Zero());
}
Run().Wait();
"#,
        ["0"]
    };

    value_task_negative_int => {
        r#"
async System.Threading.Tasks.ValueTask<int> Neg() { return -15; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Neg());
}
Run().Wait();
"#,
        ["-15"]
    };

    value_task_large_positive_int => {
        r#"
async System.Threading.Tasks.ValueTask<int> Big() { return 99999; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Big());
}
Run().Wait();
"#,
        ["99999"]
    };

    value_task_char_code_unit => {
        r#"
async System.Threading.Tasks.ValueTask<char> Get() { return 'Z'; }
async System.Threading.Tasks.Task Run() {
    char c = await Get();
    Console.WriteLine((int)c);
}
Run().Wait();
"#,
        ["90"]
    };

    value_task_expression_bodied_async => {
        r#"
async System.Threading.Tasks.ValueTask<int> Get() => 55;
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Get());
}
Run().Wait();
"#,
        ["55"]
    };

    value_task_local_async_function => {
        r#"
async System.Threading.Tasks.Task Run() {
    async System.Threading.Tasks.ValueTask<int> Local() { return 12; }
    Console.WriteLine(await Local());
}
Run().Wait();
"#,
        ["12"]
    };

    generic_value_task_method_count => {
        r#"
async System.Threading.Tasks.ValueTask<T> Identity<T>(T value) { return value; }
async System.Threading.Tasks.Task Run() {
    int count = await Identity(4);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    value_task_if_branch_true_path_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Pick(bool flag) {
    if (flag) return 100;
    return 0;
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Pick(true));
}
Run().Wait();
"#,
        ["100"]
    };

    value_task_if_branch_false_path_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Pick(bool flag) {
    if (flag) return 100;
    return 7;
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Pick(false));
}
Run().Wait();
"#,
        ["7"]
    };

    value_task_ternary_selection => {
        r#"
async System.Threading.Tasks.ValueTask<int> Choose(int a, int b, bool first) {
    return first ? a : b;
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Choose(3, 8, false));
}
Run().Wait();
"#,
        ["8"]
    };

    value_task_try_catch_recovers_with_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Risky(bool fail) {
    if (fail) throw new System.Exception("no");
    return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    try { count = await Risky(true); }
    catch (System.Exception) { count = 2; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    value_task_exception_message_length_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Fail() {
    throw new System.Exception("err");
}
async System.Threading.Tasks.Task Run() {
    int len = 0;
    try { await Fail(); }
    catch (System.Exception ex) { len = ex.Message.Length; }
    Console.WriteLine(len);
}
Run().Wait();
"#,
        ["3"]
    };

    value_task_await_same_method_twice => {
        r#"
async System.Threading.Tasks.ValueTask<int> Constant() { return 6; }
async System.Threading.Tasks.Task Run() {
    int a = await Constant();
    int b = await Constant();
    Console.WriteLine(a + b);
}
Run().Wait();
"#,
        ["12"]
    };

    value_task_chain_two_methods => {
        r#"
async System.Threading.Tasks.ValueTask<int> A() { return 2; }
async System.Threading.Tasks.ValueTask<int> B(int x) { return x + 3; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await B(await A()));
}
Run().Wait();
"#,
        ["5"]
    };

    value_task_from_task_from_result => {
        r#"
async System.Threading.Tasks.ValueTask<int> ViaTask() {
    return await System.Threading.Tasks.Task.FromResult(21);
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await ViaTask());
}
Run().Wait();
"#,
        ["21"]
    };

    value_task_configure_await_false => {
        r#"
async System.Threading.Tasks.ValueTask<int> Compute() {
    await System.Threading.Tasks.Task.Yield().ConfigureAwait(false);
    return 33;
}
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Compute());
}
Run().Wait();
"#,
        ["33"]
    };

    value_task_delay_via_as_task_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Delayed() {
    await System.Threading.Tasks.Task.Delay(0).ConfigureAwait(false);
    return 2;
}
async System.Threading.Tasks.Task Run() {
    var task = Delayed().AsTask();
    Console.WriteLine(await task);
}
Run().Wait();
"#,
        ["2"]
    };

    value_task_bool_false_count => {
        r#"
async System.Threading.Tasks.ValueTask<bool> No() { return false; }
async System.Threading.Tasks.Task Run() {
    bool v = await No();
    Console.WriteLine(v ? 1 : 0);
}
Run().Wait();
"#,
        ["0"]
    };

    value_task_empty_string_length_zero => {
        r#"
async System.Threading.Tasks.ValueTask<string> Empty() { return ""; }
async System.Threading.Tasks.Task Run() {
    string s = await Empty();
    Console.WriteLine(s.Length);
}
Run().Wait();
"#,
        ["0"]
    };

    value_task_for_accumulate_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> One() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    for (int i = 0; i < 8; i++) count += await One();
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["8"]
    };

    value_task_non_generic_void_like_count => {
        r#"
async System.Threading.Tasks.ValueTask DoWork() { return default; }
async System.Threading.Tasks.Task Run() {
    await DoWork();
    Console.WriteLine(1);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_multiply_two_results => {
        r#"
async System.Threading.Tasks.ValueTask<int> Left() { return 6; }
async System.Threading.Tasks.ValueTask<int> Right() { return 7; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Left() * await Right());
}
Run().Wait();
"#,
        ["42"]
    };

    value_task_modulo_result => {
        r#"
async System.Threading.Tasks.ValueTask<int> Dividend() { return 17; }
async System.Threading.Tasks.ValueTask<int> Divisor() { return 5; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Dividend() % await Divisor());
}
Run().Wait();
"#,
        ["2"]
    };

    value_task_bitwise_and_result => {
        r#"
async System.Threading.Tasks.ValueTask<int> Mask() { return 0xF0; }
async System.Threading.Tasks.ValueTask<int> Value() { return 0xFF; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Mask() & await Value());
}
Run().Wait();
"#,
        ["240"]
    };

    value_task_comparison_yields_count_one => {
        r#"
async System.Threading.Tasks.ValueTask<int> A() { return 10; }
async System.Threading.Tasks.ValueTask<int> B() { return 5; }
async System.Threading.Tasks.Task Run() {
    int count = (await A() > await B()) ? 1 : 0;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_null_string_length_zero => {
        r#"
async System.Threading.Tasks.ValueTask<string> NullStr() { return null; }
async System.Threading.Tasks.Task Run() {
    string s = await NullStr();
    Console.WriteLine(s == null ? 0 : s.Length);
}
Run().Wait();
"#,
        ["0"]
    };

    value_task_array_length_after_await => {
        r#"
async System.Threading.Tasks.ValueTask<int[]> Get() {
    return new int[] { 1, 2, 3, 4, 5 };
}
async System.Threading.Tasks.Task Run() {
    int[] arr = await Get();
    Console.WriteLine(arr.Length);
}
Run().Wait();
"#,
        ["5"]
    };

    value_task_list_count_property => {
        r#"
async System.Threading.Tasks.ValueTask<System.Collections.Generic.List<int>> Get() {
    return new System.Collections.Generic.List<int> { 1, 2, 3 };
}
async System.Threading.Tasks.Task Run() {
    var list = await Get();
    Console.WriteLine(list.Count);
}
Run().Wait();
"#,
        ["3"]
    };

    value_task_twice_nested_depth => {
        r#"
async System.Threading.Tasks.ValueTask<int> Deep() { return 1; }
async System.Threading.Tasks.ValueTask<int> Mid() { return await Deep() + 1; }
async System.Threading.Tasks.ValueTask<int> Top() { return await Mid() + 1; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Top());
}
Run().Wait();
"#,
        ["3"]
    };

    value_task_switch_case_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Code() { return 2; }
async System.Threading.Tasks.Task Run() {
    int c = await Code();
    int count = 0;
    switch (c) {
        case 1: count = 10; break;
        case 2: count = 20; break;
        default: count = 0; break;
    }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["20"]
    };

    value_task_while_decrement_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int n = 4;
    int count = 0;
    while (n > 0) { count += await Step(); n--; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    value_task_do_while_runs_once_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Step() { return 1; }
async System.Threading.Tasks.Task Run() {
    int count = 0;
    do { count += await Step(); } while (false);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_three_way_sum_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> N(int v) { return v; }
async System.Threading.Tasks.Task Run() {
    int count = await N(4) + await N(5) + await N(6);
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["15"]
    };

    value_task_as_task_bool_result => {
        r#"
async System.Threading.Tasks.ValueTask<bool> Yes() { return true; }
async System.Threading.Tasks.Task Run() {
    bool v = await Yes().AsTask();
    Console.WriteLine(v ? 1 : 0);
}
Run().Wait();
"#,
        ["1"]
    };

    value_task_from_result_string_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    var vt = System.Threading.Tasks.ValueTask.FromResult("abc");
    string s = await vt;
    Console.WriteLine(s.Length);
}
Run().Wait();
"#,
        ["3"]
    };

    value_task_subtraction_chain_count => {
        r#"
async System.Threading.Tasks.ValueTask<int> Start() { return 50; }
async System.Threading.Tasks.ValueTask<int> Take() { return 8; }
async System.Threading.Tasks.Task Run() {
    Console.WriteLine(await Start() - await Take());
}
Run().Wait();
"#,
        ["42"]
    };

    value_task_is_completed_before_await_sync => {
        r#"
async System.Threading.Tasks.ValueTask<int> Sync() { return 11; }
async System.Threading.Tasks.Task Run() {
    var vt = Sync();
    int count = vt.IsCompleted ? 1 : 0;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };
}
