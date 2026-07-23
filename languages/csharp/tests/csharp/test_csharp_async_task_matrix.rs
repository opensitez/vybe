use super::helpers::run_csharp;

#[test]
fn task_from_result_returns_expected_value_per_type() {
    let cases = [
        ("System.Int32", "123", "123"),
        ("System.String", "\"hello\"", "hello"),
        ("System.Boolean", "true", "True"),
        ("System.Double", "2.5", "2.5"),
    ];

    for (type_name, source_value, expected) in cases {
        let src = format!(
            "System.Threading.Tasks.Task<{type_name}> result = System.Threading.Tasks.Task.FromResult({source_value}); Console.WriteLine(result.Result.GetType().FullName); Console.WriteLine(result.Result);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![type_name.to_string(), expected.to_string()]
        );
    }
}

#[test]
fn async_local_method_returns_same_result_via_awaiter() {
    let inputs = [-5, -1, 0, 1, 2, 9, 13];

    for value in inputs {
        let expected = value * 2;
        let src = format!(
            r#"
async System.Threading.Tasks.Task<int> DoubleAsync(int x) {{
    return x * 2;
}}

Console.WriteLine(DoubleAsync({value}).GetAwaiter().GetResult());
"#
        );
        assert_eq!(run_csharp(&src), vec![expected.to_string()]);
    }
}

#[test]
fn task_run_executes_multiple_body_paths_with_deterministic_join() {
    let values = [0, 1, 2, 3, 4, 5];

    for value in values {
        let doubled = value * 2;
        let src = format!(
            r#"
var start = {value};
var task = System.Threading.Tasks.Task.Run(() => start * 2);
Console.WriteLine(task.GetAwaiter().GetResult());
"#
        );
        assert_eq!(run_csharp(&src), vec![doubled.to_string()]);
    }
}

#[test]
fn task_when_all_collects_all_results_in_input_order() {
    let triplets = [(1, 2, 3), (10, -5, 7), (-1, 0, 1), (42, 42, 42)];

    for (first, second, third) in triplets {
        let expected_sum = first + second + third;
        let src = format!(
            r#"
System.Threading.Tasks.Task<int> ToTask(int x) => System.Threading.Tasks.Task.FromResult(x);
var aggregate = System.Threading.Tasks.Task.WhenAll(ToTask({first}), ToTask({second}), ToTask({third})).GetAwaiter().GetResult();
Console.WriteLine(aggregate.Length);
Console.WriteLine(aggregate[0] + aggregate[1] + aggregate[2]);
"#
        );
        assert_eq!(
            run_csharp(&src),
            vec!["3".to_string(), expected_sum.to_string()]
        );
    }
}

#[test]
fn task_completed_task_reports_completion() {
    let src = r#"
var completed = System.Threading.Tasks.Task.CompletedTask;
Console.WriteLine(completed.IsCompleted);
Console.WriteLine(completed.IsFaulted);
"#;
    assert_eq!(
        run_csharp(src),
        vec!["True".to_string(), "False".to_string()]
    );
}
