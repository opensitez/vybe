//! `IAsyncEnumerable<T>` — async streams, `yield return`, and `await foreach`.

csharp_cases! {
    await_foreach_counts_three_item_stream => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield return 3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    empty_async_enumerable_produces_zero_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Empty() {
    yield break;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Empty()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["0"]
    };

    single_yield_async_stream_count_is_one => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> One() {
    yield return 99;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in One()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["1"]
    };

    await_foreach_sums_async_stream => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int s = 0;
    await foreach (var x in Stream()) s += x;
    Console.WriteLine(s);
}
Run().Wait();
"#,
        ["3"]
    };

    task_yield_between_yields_preserves_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 10;
    await System.Threading.Tasks.Task.Yield();
    yield return 20;
    await System.Threading.Tasks.Task.Yield();
    yield return 30;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    yield_break_truncates_async_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield break;
    yield return 99;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    nested_await_foreach_counts_both_streams => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Inner() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> Outer() {
    await foreach (var x in Inner()) yield return x;
    await foreach (var x in Inner()) yield return x;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Outer()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    async_enumerable_string_items_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<string> Words() {
    yield return "a";
    yield return "bb";
    yield return "ccc";
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var w in Words()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    async_enumerable_bool_true_false_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<bool> Flags() {
    yield return true;
    yield return false;
    yield return true;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var f in Flags()) if (f) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    async_local_function_stream_count => {
        r#"
async System.Threading.Tasks.Task Run() {
    async System.Collections.Generic.IAsyncEnumerable<int> Local() {
        for (int i = 0; i < 5; i++) yield return i;
    }
    int count = 0;
    await foreach (var x in Local()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    generic_async_enumerable_method_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<T> Repeat<T>(T value, int times) {
    for (int i = 0; i < times; i++) yield return value;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Repeat(7, 4)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    cancellation_token_default_completes_full_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken = default) {
    for (int i = 0; i < 6; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    cancellation_token_none_runs_entire_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 0; i < 4; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream(System.Threading.CancellationToken.None)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    cancellation_token_cancelled_before_foreach_yields_zero => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 0; i < 8; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    cts.Cancel();
    int count = 0;
    try {
        await foreach (var x in Stream(cts.Token)) count++;
    } catch (System.OperationCanceledException) { }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["0"]
    };

    explicit_cancellation_token_passed_counts_all_items => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken cancellationToken) {
    for (int i = 1; i <= 5; i++) {
        cancellationToken.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    await foreach (var x in Stream(cts.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    break_from_await_foreach_stops_early_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 10; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        count++;
        if (x == 2) break;
    }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    continue_in_await_foreach_skips_matching_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 6; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        if (x % 2 == 0) continue;
        count++;
    }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    await_foreach_tracks_index_and_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 5;
    yield return 10;
    yield return 15;
}
async System.Threading.Tasks.Task Run() {
    int index = 0;
    await foreach (var x in Stream()) index++;
    Console.WriteLine(index);
}
Run().Wait();
"#,
        ["3"]
    };

    conditional_yield_only_evens_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Evens(int max) {
    for (int i = 0; i < max; i++)
        if (i % 2 == 0) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Evens(7)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    async_stream_negative_numbers_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Negatives() {
    yield return -1;
    yield return -2;
    yield return -3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Negatives()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    async_stream_all_zeros_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Zeros() {
    yield return 0;
    yield return 0;
    yield return 0;
    yield return 0;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Zeros()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    async_stream_loop_twenty_items_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Range20() {
    for (int i = 0; i < 20; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Range20()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["20"]
    };

    two_digit_values_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Tens() {
    yield return 10;
    yield return 20;
    yield return 30;
    yield return 40;
    yield return 50;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Tens()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    sequential_await_foreach_on_fresh_factory_counts_twice => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int total = 0;
    await foreach (var x in Make()) total++;
    await foreach (var x in Make()) total++;
    Console.WriteLine(total);
}
Run().Wait();
"#,
        ["4"]
    };

    interface_variable_async_enumerable_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Make() {
    yield return 3;
    yield return 6;
    yield return 9;
}
async System.Threading.Tasks.Task Run() {
    System.Collections.Generic.IAsyncEnumerable<int> stream = Make();
    int count = 0;
    await foreach (var x in stream) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    instance_method_async_stream_count => {
        r#"
class Counter {
    public async System.Collections.Generic.IAsyncEnumerable<int> Stream(int n) {
        for (int i = 0; i < n; i++) yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var c = new Counter();
    int count = 0;
    await foreach (var x in c.Stream(7)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["7"]
    };

    static_async_enumerable_method_count => {
        r#"
class Factory {
    public static async System.Collections.Generic.IAsyncEnumerable<int> Three() {
        yield return 1;
        yield return 2;
        yield return 3;
    }
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Factory.Three()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    explicit_int_type_in_await_foreach_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 4;
    yield return 8;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (int x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    await_foreach_product_nonzero_count_check => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int product = 1;
    int count = 0;
    await foreach (var x in Stream()) { product *= x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    await_foreach_running_max_count_at_end => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 3;
    yield return 7;
    yield return 5;
    yield return 9;
}
async System.Threading.Tasks.Task Run() {
    int max = int.MinValue;
    int count = 0;
    await foreach (var x in Stream()) { if (x > max) max = x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    odd_only_conditional_yield_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Odds(int n) {
    for (int i = 0; i < n; i++)
        if (i % 2 == 1) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Odds(8)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    skip_first_two_via_flag_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 0; i < 6; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int seen = 0;
    int count = 0;
    await foreach (var x in Stream()) {
        seen++;
        if (seen <= 2) continue;
        count++;
    }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    take_until_threshold_reached_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    for (int i = 1; i <= 10; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) {
        count++;
        if (x >= 5) break;
    }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    nested_loop_yield_flat_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Grid() {
    for (int r = 0; r < 2; r++)
        for (int c = 0; c < 3; c++)
            yield return r * 10 + c;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Grid()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    async_enumerable_char_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<char> Chars() {
    yield return 'a';
    yield return 'b';
    yield return 'c';
    yield return 'd';
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var ch in Chars()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    async_enumerable_byte_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<byte> Bytes() {
    yield return (byte)1;
    yield return (byte)2;
    yield return (byte)3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var b in Bytes()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    async_enumerable_long_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<long> Longs() {
    yield return 100L;
    yield return 200L;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var v in Longs()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    delay_zero_between_yields_keeps_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    await System.Threading.Tasks.Task.Delay(0);
    yield return 2;
    await System.Threading.Tasks.Task.Delay(0);
    yield return 3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    immediate_yield_break_empty_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Nothing() {
    yield break;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Nothing()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["0"]
    };

    filter_positive_values_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Mixed() {
    yield return -1;
    yield return 2;
    yield return -3;
    yield return 4;
    yield return 5;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Mixed()) if (x > 0) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    map_in_foreach_body_count_unchanged => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 1;
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    int sum = 0;
    await foreach (var x in Stream()) { sum += x * 2; count++; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    two_streams_sequential_count_total => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> A() {
    yield return 1;
    yield return 2;
}
async System.Collections.Generic.IAsyncEnumerable<int> B() {
    yield return 10;
    yield return 20;
    yield return 30;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in A()) count++;
    await foreach (var x in B()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    range_style_async_enumerable_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> FromTo(int start, int end) {
    for (int i = start; i <= end; i++) yield return i;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in FromTo(3, 8)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["6"]
    };

    token_passed_but_not_cancelled_full_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 9; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    await foreach (var x in Stream(cts.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["9"]
    };

    cancellation_after_three_yields_partial_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 10; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var cts = new System.Threading.CancellationTokenSource();
    int count = 0;
    try {
        await foreach (var x in Stream(cts.Token)) {
            count++;
            if (count == 3) cts.Cancel();
        }
    } catch (System.OperationCanceledException) { }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    independent_second_stream_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Pair() {
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Pair()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    factory_called_twice_yields_fresh_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Fresh() {
    yield return 100;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Fresh()) count++;
    await foreach (var x in Fresh()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    params_style_loop_yield_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> FromValues(int a, int b, int c, int d) {
    yield return a;
    yield return b;
    yield return c;
    yield return d;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in FromValues(1, 2, 3, 4)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["4"]
    };

    nullable_int_stream_non_null_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int?> Stream() {
    yield return 1;
    yield return null;
    yield return 3;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Stream()) if (x != null) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    square_accumulator_with_count_output => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream() {
    yield return 2;
    yield return 3;
    yield return 4;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    int sumSq = 0;
    await foreach (var x in Stream()) { sumSq += x * x; count++; }
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["3"]
    };

    async_stream_with_await_before_first_yield_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> DelayedStart() {
    await System.Threading.Tasks.Task.Yield();
    yield return 1;
    yield return 2;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in DelayedStart()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    linked_cancellation_token_source_full_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 5; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var parent = new System.Threading.CancellationTokenSource();
    var linked = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(parent.Token);
    int count = 0;
    await foreach (var x in Stream(linked.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["5"]
    };

    await_foreach_string_length_total_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<string> Stream() {
    yield return "ab";
    yield return "cde";
    yield return "f";
}
async System.Threading.Tasks.Task Run() {
    int totalLen = 0;
    await foreach (var s in Stream()) totalLen += s.Length;
    Console.WriteLine(totalLen);
}
Run().Wait();
"#,
        ["6"]
    };

    double_nested_async_enumerable_flat_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Inner() {
    yield return 1;
}
async System.Collections.Generic.IAsyncEnumerable<int> Middle() {
    await foreach (var x in Inner()) yield return x;
}
async System.Collections.Generic.IAsyncEnumerable<int> Outer() {
    await foreach (var x in Middle()) yield return x;
    await foreach (var x in Middle()) yield return x;
}
async System.Threading.Tasks.Task Run() {
    int count = 0;
    await foreach (var x in Outer()) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["2"]
    };

    cancellation_token_linked_parent_not_cancelled_count => {
        r#"
async System.Collections.Generic.IAsyncEnumerable<int> Stream(
    System.Threading.CancellationToken token) {
    for (int i = 0; i < 7; i++) {
        token.ThrowIfCancellationRequested();
        yield return i;
    }
}
async System.Threading.Tasks.Task Run() {
    var parent = new System.Threading.CancellationTokenSource();
    var child = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(parent.Token);
    int count = 0;
    await foreach (var x in Stream(child.Token)) count++;
    Console.WriteLine(count);
}
Run().Wait();
"#,
        ["7"]
    };
}
