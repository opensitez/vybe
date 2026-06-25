//! `async`/`await` resumes after nested task completes with observed result.
use super::helpers::run_csharp;

#[test]
fn await_on_task_from_result_returns_completed_value_to_caller() {
    assert_eq!(
        run_csharp(
            r#"
using System.Threading.Tasks;
async Task<int> Load() { return await Task.FromResult(9); }
async Task Run() { Console.WriteLine(await Load()); }
Run().GetAwaiter().GetResult();
"#
        ),
        &["9"]
    );
}

#[test]
fn await_chains_through_two_async_methods_preserving_return_type() {
    assert_eq!(
        run_csharp(
            r#"
using System.Threading.Tasks;
async Task<int> Inner() { return await Task.FromResult(4); }
async Task<int> Outer() { return await Inner() + 1; }
Console.WriteLine(Outer().GetAwaiter().GetResult());
"#
        ),
        &["5"]
    );
}

#[test]
fn async_void_local_function_can_be_awaited_from_enclosing_async_method() {
    assert_eq!(
        run_csharp(
            r#"
using System.Threading.Tasks;
async Task Run() {
    async Task<int> Compute() { return await Task.FromResult(6); }
    Console.WriteLine(await Compute());
}
Run().GetAwaiter().GetResult();
"#
        ),
        &["6"]
    );
}

#[test]
fn await_in_try_finally_still_runs_finally_before_result_is_observed() {
    assert_eq!(
        run_csharp(
            r#"
using System.Threading.Tasks;
async Task<int> Pick() {
    try {
        return await Task.FromResult(2);
    } finally {
        Console.WriteLine("cleanup");
    }
}
Console.WriteLine(Pick().GetAwaiter().GetResult());
"#
        ),
        &["cleanup", "2"]
    );
}

#[test]
fn task_run_offloads_work_and_returns_result_to_awaiter() {
    assert_eq!(
        run_csharp(
            r#"
using System.Threading.Tasks;
async Task<int> Run() {
    return await Task.Run(() => 11);
}
Console.WriteLine(Run().GetAwaiter().GetResult());
"#
        ),
        &["11"]
    );
}
