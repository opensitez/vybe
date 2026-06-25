//! Threading and synchronization: `Lazy<T>`, `Interlocked`, `[ThreadStatic]`, `WeakReference`.
use super::helpers::run_csharp;

#[test]
fn lazy_factory_runs_once_on_first_value_access() {
    assert_eq!(
        run_csharp(
            r#"
int calls = 0;
var lazy = new System.Lazy<int>(() => { calls++; return 7; });
Console.WriteLine(calls);
Console.WriteLine(lazy.Value);
Console.WriteLine(calls);
"#
        ),
        &["0", "7", "1"]
    );
}

#[test]
fn lazy_is_value_created_flips_after_materialization() {
    assert_eq!(
        run_csharp(
            r#"
var lazy = new System.Lazy<int>(() => 3);
Console.WriteLine(lazy.IsValueCreated);
Console.WriteLine(lazy.Value);
Console.WriteLine(lazy.IsValueCreated);
"#
        ),
        &["False", "3", "True"]
    );
}

#[test]
fn interlocked_increment_atomically_adds_one() {
    assert_eq!(
        run_csharp(
            r#"
int count = 5;
Console.WriteLine(System.Threading.Interlocked.Increment(ref count));
Console.WriteLine(count);
"#
        ),
        &["6", "6"]
    );
}

#[test]
fn interlocked_add_returns_sum_and_updates_storage() {
    assert_eq!(
        run_csharp(
            r#"
int total = 10;
Console.WriteLine(System.Threading.Interlocked.Add(ref total, 4));
Console.WriteLine(total);
"#
        ),
        &["14", "14"]
    );
}

#[test]
fn interlocked_exchange_swaps_value_and_returns_previous() {
    assert_eq!(
        run_csharp(
            r#"
int slot = 1;
Console.WriteLine(System.Threading.Interlocked.Exchange(ref slot, 9));
Console.WriteLine(slot);
"#
        ),
        &["1", "9"]
    );
}

#[test]
fn interlocked_compare_exchange_updates_only_when_current_matches() {
    assert_eq!(
        run_csharp(
            r#"
int slot = 7;
var previous = System.Threading.Interlocked.CompareExchange(ref slot, 99, 7);
Console.WriteLine(previous);
Console.WriteLine(slot);
"#
        ),
        &["7", "99"]
    );
}

#[test]
fn thread_static_field_defaults_to_zero_on_main_thread() {
    assert_eq!(
        run_csharp(
            r#"
class Counter {
    [System.ThreadStatic]
    public static int Value;
}
Console.WriteLine(Counter.Value);
"#
        ),
        &["0"]
    );
}

#[test]
fn weak_reference_target_is_alive_while_strong_reference_exists() {
    assert_eq!(
        run_csharp(
            r#"
var strong = new object();
var weak = new System.WeakReference(strong);
Console.WriteLine(weak.IsAlive);
"#
        ),
        &["True"]
    );
}

#[test]
fn weak_reference_try_get_target_returns_false_after_target_collected() {
    assert_eq!(
        run_csharp(
            r#"
System.WeakReference weak;
void Create() {
    weak = new System.WeakReference(new object());
}
Create();
System.GC.Collect();
System.GC.WaitForPendingFinalizers();
Console.WriteLine(weak.TryGetTarget(out var target));
"#
        ),
        &["False"]
    );
}
