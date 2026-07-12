//! `Interlocked` atomic operations — `Increment`, `Add`, `CompareExchange`, `Exchange`.
//! GAP: concurrency primitives.

csharp_cases! {
    interlocked_increment_from_zero_count => {
        r#"
int counter = 0;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["1", "1"]
    };

    interlocked_increment_from_five_count => {
        r#"
int counter = 5;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["6", "6"]
    };

    interlocked_increment_from_negative_one => {
        r#"
int counter = -1;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["0", "0"]
    };

    interlocked_decrement_from_ten_count => {
        r#"
int counter = 10;
Console.WriteLine(System.Threading.Interlocked.Decrement(ref counter));
Console.WriteLine(counter);
"#,
        ["9", "9"]
    };

    interlocked_add_positive_delta => {
        r#"
int total = 10;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, 4));
Console.WriteLine(total);
"#,
        ["14", "14"]
    };

    interlocked_add_negative_delta => {
        r#"
int total = 20;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, -5));
Console.WriteLine(total);
"#,
        ["15", "15"]
    };

    interlocked_add_zero_leaves_value => {
        r#"
int total = 8;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, 0));
Console.WriteLine(total);
"#,
        ["8", "8"]
    };

    interlocked_exchange_returns_old_and_sets_new => {
        r#"
int slot = 1;
Console.WriteLine(System.Threading.Interlocked.Exchange(ref slot, 9));
Console.WriteLine(slot);
"#,
        ["1", "9"]
    };

    interlocked_exchange_from_zero => {
        r#"
int slot = 0;
Console.WriteLine(System.Threading.Interlocked.Exchange(ref slot, 42));
Console.WriteLine(slot);
"#,
        ["0", "42"]
    };

    interlocked_exchange_overwrites_existing => {
        r#"
int slot = 77;
Console.WriteLine(System.Threading.Interlocked.Exchange(ref slot, 3));
Console.WriteLine(slot);
"#,
        ["77", "3"]
    };

    interlocked_compare_exchange_match_updates => {
        r#"
int slot = 7;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 99, 7);
Console.WriteLine(previous);
Console.WriteLine(slot);
"#,
        ["7", "99"]
    };

    interlocked_compare_exchange_no_match_keeps_old => {
        r#"
int slot = 7;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 99, 8);
Console.WriteLine(previous);
Console.WriteLine(slot);
"#,
        ["7", "7"]
    };

    interlocked_compare_exchange_from_zero => {
        r#"
int slot = 0;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 5, 0);
Console.WriteLine(previous);
Console.WriteLine(slot);
"#,
        ["0", "5"]
    };

    interlocked_increment_twice_count => {
        r#"
int counter = 0;
System.Threading.Interlocked.Increment(ref counter);
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["1", "2", "2"]
    };

    interlocked_add_then_read_total => {
        r#"
int total = 3;
System.Threading.Interlocked.Add(ref total, 7);
Console.WriteLine(total);
"#,
        ["10"]
    };

    interlocked_exchange_then_increment => {
        r#"
int slot = 2;
System.Threading.Interlocked.Exchange(ref slot, 10);
Console.WriteLine(System.Threading.Interlocked.Increment(ref slot));
Console.WriteLine(slot);
"#,
        ["11", "11"]
    };

    interlocked_compare_exchange_then_add => {
        r#"
int slot = 4;
System.Threading.Interlocked.CompareExchange(ref slot, 4, 4);
Console.WriteLine(System.Threading.Interlocked.Add(ref slot, 6));
Console.WriteLine(slot);
"#,
        ["10", "10"]
    };

    interlocked_loop_five_increments_count => {
        r#"
int counter = 0;
for (int i = 0; i < 5; i++) System.Threading.Interlocked.Increment(ref counter);
Console.WriteLine(counter);
"#,
        ["5"]
    };

    interlocked_loop_add_accumulates => {
        r#"
int total = 0;
for (int i = 1; i <= 4; i++) System.Threading.Interlocked.Add(ref total, i);
Console.WriteLine(total);
"#,
        ["10"]
    };

    interlocked_increment_large_start => {
        r#"
int counter = 999;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["1000", "1000"]
    };

    interlocked_add_large_delta => {
        r#"
int total = 100;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, 900));
Console.WriteLine(total);
"#,
        ["1000", "1000"]
    };

    interlocked_exchange_with_zero_new => {
        r#"
int slot = 55;
Console.WriteLine(System.Threading.Interlocked.Exchange(ref slot, 0));
Console.WriteLine(slot);
"#,
        ["55", "0"]
    };

    interlocked_compare_exchange_expected_zero => {
        r#"
int slot = 0;
var prev = System.Threading.Interlocked.CompareExchange(ref slot, 12, 0);
Console.WriteLine(prev);
Console.WriteLine(slot);
"#,
        ["0", "12"]
    };

    interlocked_compare_exchange_wrong_expected => {
        r#"
int slot = 12;
var prev = System.Threading.Interlocked.CompareExchange(ref slot, 99, 0);
Console.WriteLine(prev);
Console.WriteLine(slot);
"#,
        ["12", "12"]
    };

    interlocked_decrement_to_zero => {
        r#"
int counter = 1;
Console.WriteLine(System.Threading.Interlocked.Decrement(ref counter));
Console.WriteLine(counter);
"#,
        ["0", "0"]
    };

    interlocked_decrement_below_zero => {
        r#"
int counter = 0;
Console.WriteLine(System.Threading.Interlocked.Decrement(ref counter));
Console.WriteLine(counter);
"#,
        ["-1", "-1"]
    };

    interlocked_add_subtract_net_zero => {
        r#"
int total = 50;
System.Threading.Interlocked.Add(ref total, 10);
System.Threading.Interlocked.Add(ref total, -10);
Console.WriteLine(total);
"#,
        ["50"]
    };

    interlocked_exchange_twice_final_value => {
        r#"
int slot = 1;
System.Threading.Interlocked.Exchange(ref slot, 2);
System.Threading.Interlocked.Exchange(ref slot, 3);
Console.WriteLine(slot);
"#,
        ["3"]
    };

    interlocked_increment_from_ninety_nine => {
        r#"
int counter = 99;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["100", "100"]
    };

    interlocked_add_one_equals_increment => {
        r#"
int a = 6;
int b = 6;
System.Threading.Interlocked.Increment(ref a);
System.Threading.Interlocked.Add(ref b, 1);
Console.WriteLine(a + b);
"#,
        ["14"]
    };

    interlocked_compare_exchange_idempotent_same => {
        r#"
int slot = 3;
var p1 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
var p2 = System.Threading.Interlocked.CompareExchange(ref slot, 8, 3);
Console.WriteLine(p1 + p2);
Console.WriteLine(slot);
"#,
        ["6", "8"]
    };

    interlocked_task_run_increment_count => {
        r#"
int counter = 0;
var tasks = new System.Threading.Tasks.Task[5];
for (int i = 0; i < 5; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Increment(ref counter);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["5"]
    };

    interlocked_task_run_add_count => {
        r#"
int total = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        System.Threading.Interlocked.Add(ref total, 2);
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(total);
"#,
        ["8"]
    };

    interlocked_exchange_in_loop_count => {
        r#"
int slot = 0;
for (int i = 1; i <= 3; i++) {
    System.Threading.Interlocked.Exchange(ref slot, i);
}
Console.WriteLine(slot);
"#,
        ["3"]
    };

    interlocked_compare_exchange_cas_retry_pattern => {
        r#"
int slot = 0;
int observed = slot;
int desired = observed + 1;
while (observed != System.Threading.Interlocked.CompareExchange(ref slot, desired, observed)) {
    observed = slot;
    desired = observed + 1;
}
Console.WriteLine(slot);
"#,
        ["1"]
    };

    interlocked_increment_then_decrement_net_zero => {
        r#"
int counter = 0;
System.Threading.Interlocked.Increment(ref counter);
System.Threading.Interlocked.Decrement(ref counter);
Console.WriteLine(counter);
"#,
        ["0"]
    };

    interlocked_add_multiple_steps_sum => {
        r#"
int total = 0;
System.Threading.Interlocked.Add(ref total, 2);
System.Threading.Interlocked.Add(ref total, 3);
System.Threading.Interlocked.Add(ref total, 5);
Console.WriteLine(total);
"#,
        ["10"]
    };

    interlocked_exchange_returns_each_previous => {
        r#"
int slot = 4;
int old1 = System.Threading.Interlocked.Exchange(ref slot, 5);
int old2 = System.Threading.Interlocked.Exchange(ref slot, 6);
Console.WriteLine(old1 + old2);
Console.WriteLine(slot);
"#,
        ["9", "6"]
    };

    interlocked_increment_from_minus_five => {
        r#"
int counter = -5;
Console.WriteLine(System.Threading.Interlocked.Increment(ref counter));
Console.WriteLine(counter);
"#,
        ["-4", "-4"]
    };

    interlocked_add_minus_three_to_ten => {
        r#"
int total = 10;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, -3));
Console.WriteLine(total);
"#,
        ["7", "7"]
    };

    interlocked_compare_exchange_negative_values => {
        r#"
int slot = -2;
var prev = System.Threading.Interlocked.CompareExchange(ref slot, -1, -2);
Console.WriteLine(prev);
Console.WriteLine(slot);
"#,
        ["-2", "-1"]
    };

    interlocked_loop_ten_add_one_count => {
        r#"
int counter = 0;
for (int i = 0; i < 10; i++) System.Threading.Interlocked.Add(ref counter, 1);
Console.WriteLine(counter);
"#,
        ["10"]
    };

    interlocked_task_run_exchange_count => {
        r#"
int slot = 0;
var t1 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 1));
var t2 = System.Threading.Tasks.Task.Run(() => System.Threading.Interlocked.Exchange(ref slot, 2));
System.Threading.Tasks.Task.WaitAll(t1, t2);
Console.WriteLine(slot);
"#,
        ["2"]
    };

    interlocked_increment_field_on_class => {
        r#"
class Counter {
    public int Value = 0;
    public void Bump() { System.Threading.Interlocked.Increment(ref Value); }
}
var c = new Counter();
c.Bump();
c.Bump();
Console.WriteLine(c.Value);
"#,
        ["2"]
    };

    interlocked_add_field_on_class => {
        r#"
class Counter {
    public int Value = 5;
    public void Add(int n) { System.Threading.Interlocked.Add(ref Value, n); }
}
var c = new Counter();
c.Add(3);
Console.WriteLine(c.Value);
"#,
        ["8"]
    };

    interlocked_compare_exchange_field_on_class => {
        r#"
class Counter {
    public int Value = 0;
    public int Cas(int expected, int desired) {
        return System.Threading.Interlocked.CompareExchange(ref Value, desired, expected);
    }
}
var c = new Counter();
Console.WriteLine(c.Cas(0, 11));
Console.WriteLine(c.Value);
"#,
        ["0", "11"]
    };
}
