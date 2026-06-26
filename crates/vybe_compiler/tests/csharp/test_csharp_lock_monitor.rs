//! `lock` statement and `Monitor` — mutual exclusion via shared counter prints.
//! GAP: concurrency primitives.

use crate::csharp_cases;

csharp_cases! {
    lock_single_increment_leaves_counter_one => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) { counter++; }
Console.WriteLine(counter);
"#,
        ["1"]
    };

    lock_three_sequential_increments_count => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) { counter++; }
lock (gate) { counter++; }
lock (gate) { counter++; }
Console.WriteLine(counter);
"#,
        ["3"]
    };

    lock_read_modify_write_doubles_counter => {
        r#"
object gate = new object();
int counter = 2;
lock (gate) { counter = counter * 2; }
Console.WriteLine(counter);
"#,
        ["4"]
    };

    lock_nested_reentrant_same_object_count => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) {
    counter++;
    lock (gate) { counter++; }
}
Console.WriteLine(counter);
"#,
        ["2"]
    };

    lock_preserves_counter_when_no_contention => {
        r#"
object gate = new object();
int counter = 7;
lock (gate) { counter += 3; }
Console.WriteLine(counter);
"#,
        ["10"]
    };

    lock_body_assigns_direct_count => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) { counter = 15; }
Console.WriteLine(counter);
"#,
        ["15"]
    };

    lock_loop_ten_times_count => {
        r#"
object gate = new object();
int counter = 0;
for (int i = 0; i < 10; i++) lock (gate) { counter++; }
Console.WriteLine(counter);
"#,
        ["10"]
    };

    lock_on_this_reference_increments_field => {
        r#"
class Box {
    public int counter = 0;
    public void Inc() { lock (this) { counter++; } }
}
var box = new Box();
box.Inc();
Console.WriteLine(box.counter);
"#,
        ["1"]
    };

    lock_two_objects_independent_counters => {
        r#"
object a = new object();
object b = new object();
int ca = 0;
int cb = 0;
lock (a) { ca++; }
lock (b) { cb += 2; }
Console.WriteLine(ca + cb);
"#,
        ["3"]
    };

    monitor_enter_exit_increments_once => {
        r#"
object gate = new object();
int counter = 0;
System.Threading.Monitor.Enter(gate);
counter++;
System.Threading.Monitor.Exit(gate);
Console.WriteLine(counter);
"#,
        ["1"]
    };

    monitor_try_enter_succeeds_when_unlocked => {
        r#"
object gate = new object();
bool got = System.Threading.Monitor.TryEnter(gate, 0);
if (got) System.Threading.Monitor.Exit(gate);
Console.WriteLine(got ? 1 : 0);
"#,
        ["1"]
    };

    monitor_try_enter_fails_when_already_locked => {
        r#"
object gate = new object();
System.Threading.Monitor.Enter(gate);
bool got = System.Threading.Monitor.TryEnter(gate, 0);
System.Threading.Monitor.Exit(gate);
Console.WriteLine(got ? 1 : 0);
"#,
        ["0"]
    };

    lock_task_run_two_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[2];
for (int i = 0; i < 2; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["2"]
    };

    lock_task_run_three_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[3];
for (int i = 0; i < 3; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["3"]
    };

    lock_task_run_five_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[5];
for (int i = 0; i < 5; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["5"]
    };

    lock_task_run_adds_two_per_worker_count => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter += 2; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["8"]
    };

    lock_without_contention_read_then_write => {
        r#"
object gate = new object();
int counter = 1;
lock (gate) {
    int snapshot = counter;
    counter = snapshot + 4;
}
Console.WriteLine(counter);
"#,
        ["5"]
    };

    lock_decrement_inside_body_count => {
        r#"
object gate = new object();
int counter = 5;
lock (gate) { counter--; }
Console.WriteLine(counter);
"#,
        ["4"]
    };

    lock_multiple_assignments_last_wins => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) {
    counter = 1;
    counter = 2;
    counter = 9;
}
Console.WriteLine(counter);
"#,
        ["9"]
    };

    lock_local_function_increments_counter => {
        r#"
object gate = new object();
int counter = 0;
void Bump() { lock (gate) { counter++; } }
Bump();
Bump();
Console.WriteLine(counter);
"#,
        ["2"]
    };

    monitor_is_entered_true_while_holding_lock => {
        r#"
object gate = new object();
int count = 0;
System.Threading.Monitor.Enter(gate);
count = System.Threading.Monitor.IsEntered(gate) ? 1 : 0;
System.Threading.Monitor.Exit(gate);
Console.WriteLine(count);
"#,
        ["1"]
    };

    monitor_is_entered_false_after_exit => {
        r#"
object gate = new object();
System.Threading.Monitor.Enter(gate);
System.Threading.Monitor.Exit(gate);
Console.WriteLine(System.Threading.Monitor.IsEntered(gate) ? 1 : 0);
"#,
        ["0"]
    };

    lock_task_run_eight_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[8];
for (int i = 0; i < 8; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["8"]
    };

    lock_separate_gates_no_interference => {
        r#"
object g1 = new object();
object g2 = new object();
int c1 = 0;
int c2 = 0;
lock (g1) { c1 = 3; }
lock (g2) { c2 = 4; }
Console.WriteLine(c1 + c2);
"#,
        ["7"]
    };

    lock_while_loop_increments_to_limit => {
        r#"
object gate = new object();
int counter = 0;
int n = 6;
while (n > 0) {
    lock (gate) { counter++; }
    n--;
}
Console.WriteLine(counter);
"#,
        ["6"]
    };

    lock_if_branch_increments_once => {
        r#"
object gate = new object();
int counter = 0;
bool flag = true;
lock (gate) { if (flag) counter++; }
Console.WriteLine(counter);
"#,
        ["1"]
    };

    lock_if_branch_skips_increment => {
        r#"
object gate = new object();
int counter = 0;
bool flag = false;
lock (gate) { if (flag) counter++; }
Console.WriteLine(counter);
"#,
        ["0"]
    };

    lock_task_run_ten_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[10];
for (int i = 0; i < 10; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["10"]
    };

    lock_triple_nested_reentrant_count => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) {
    counter++;
    lock (gate) {
        counter++;
        lock (gate) { counter++; }
    }
}
Console.WriteLine(counter);
"#,
        ["3"]
    };

    lock_monitor_mixed_enter_and_lock_count => {
        r#"
object gate = new object();
int counter = 0;
System.Threading.Monitor.Enter(gate);
counter++;
lock (gate) { counter++; }
System.Threading.Monitor.Exit(gate);
Console.WriteLine(counter);
"#,
        ["2"]
    };

    lock_task_run_four_workers_triple_add => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[4];
for (int i = 0; i < 4; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter += 3; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["12"]
    };

    lock_zero_initial_increment_count => {
        r#"
object gate = new object();
int counter = 0;
lock (gate) { counter = counter + 0 + 1; }
Console.WriteLine(counter);
"#,
        ["1"]
    };

    lock_negative_counter_increment => {
        r#"
object gate = new object();
int counter = -2;
lock (gate) { counter++; }
Console.WriteLine(counter);
"#,
        ["-1"]
    };

    lock_large_counter_addition => {
        r#"
object gate = new object();
int counter = 1000;
lock (gate) { counter += 250; }
Console.WriteLine(counter);
"#,
        ["1250"]
    };

    lock_task_run_six_workers_double_increment => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[6];
for (int i = 0; i < 6; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => {
        lock (gate) { counter++; counter++; }
    });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["12"]
    };

    lock_on_string_interned_gate_count => {
        r#"
string gate = "sync-root";
int counter = 0;
lock (gate) { counter++; }
Console.WriteLine(counter);
"#,
        ["1"]
    };

    lock_do_while_once_increments => {
        r#"
object gate = new object();
int counter = 0;
do { lock (gate) { counter++; } } while (false);
Console.WriteLine(counter);
"#,
        ["1"]
    };

    lock_for_loop_accumulates_squares_count => {
        r#"
object gate = new object();
int counter = 0;
for (int i = 1; i <= 4; i++) {
    lock (gate) { counter += i * i; }
}
Console.WriteLine(counter);
"#,
        ["30"]
    };

    monitor_try_enter_timeout_zero_unlocked_count => {
        r#"
object gate = new object();
bool got = System.Threading.Monitor.TryEnter(gate);
int count = got ? 1 : 0;
if (got) System.Threading.Monitor.Exit(gate);
Console.WriteLine(count);
"#,
        ["1"]
    };

    lock_class_instance_field_gate => {
        r#"
class SafeCounter {
    private object gate = new object();
    public int Value = 0;
    public void Add(int n) { lock (gate) { Value += n; } }
}
var sc = new SafeCounter();
sc.Add(3);
sc.Add(4);
Console.WriteLine(sc.Value);
"#,
        ["7"]
    };

    lock_task_run_two_gates_isolated_totals => {
        r#"
object g1 = new object();
object g2 = new object();
int c1 = 0;
int c2 = 0;
var t1 = System.Threading.Tasks.Task.Run(() => { lock (g1) { c1 += 5; } });
var t2 = System.Threading.Tasks.Task.Run(() => { lock (g2) { c2 += 6; } });
System.Threading.Tasks.Task.WaitAll(t1, t2);
Console.WriteLine(c1 + c2);
"#,
        ["11"]
    };

    lock_switch_case_increments_matching => {
        r#"
object gate = new object();
int counter = 0;
int code = 2;
lock (gate) {
    switch (code) {
        case 1: counter = 10; break;
        case 2: counter = 20; break;
        default: counter = 0; break;
    }
}
Console.WriteLine(counter);
"#,
        ["20"]
    };

    lock_ternary_assignment_count => {
        r#"
object gate = new object();
int counter = 0;
bool pick = true;
lock (gate) { counter = pick ? 3 : 8; }
Console.WriteLine(counter);
"#,
        ["3"]
    };

    lock_task_run_seven_workers_counter => {
        r#"
object gate = new object();
int counter = 0;
var tasks = new System.Threading.Tasks.Task[7];
for (int i = 0; i < 7; i++) {
    tasks[i] = System.Threading.Tasks.Task.Run(() => { lock (gate) { counter++; } });
}
System.Threading.Tasks.Task.WaitAll(tasks);
Console.WriteLine(counter);
"#,
        ["7"]
    };

    lock_nested_different_objects_two_counts => {
        r#"
object outer = new object();
object inner = new object();
int counter = 0;
lock (outer) {
    counter++;
    lock (inner) { counter++; }
}
Console.WriteLine(counter);
"#,
        ["2"]
    };
}
