use super::helpers::run_csharp;
use std::collections::HashSet;

fn bool_text(v: bool) -> &'static str {
    if v { "True" } else { "False" }
}

#[test]
fn list_matrix_sum_count_first_and_contains_even() {
    let cases: Vec<Vec<i32>> = vec![vec![1, 2, 3, 4], vec![-2, 0, 7], vec![9], vec![]];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("list.Add({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let src = format!(
            "using System.Collections.Generic; var list = new List<int>(); {add_lines} int sum = 0; foreach (var value in list) {{ sum += value; }} Console.WriteLine(sum); Console.WriteLine(list.Count); Console.WriteLine(list.Count == 0 ? 0 : list[0]); Console.WriteLine(list.Contains(2));"
        );
        let expected_sum: i32 = values.iter().sum();
        let expected_first = values.first().copied().unwrap_or(0);
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_sum.to_string(),
                values.len().to_string(),
                expected_first.to_string(),
                bool_text(values.contains(&2)).to_string(),
            ]
        );
    }
}

#[test]
fn list_matrix_mutate_front_and_report_head() {
    let cases: Vec<Vec<i32>> = vec![vec![3, 2, 1], vec![10, 20, 30, 40], vec![7], vec![]];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("list.Add({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let src = format!(
            "using System.Collections.Generic; var list = new List<int>(); {add_lines} if (list.Count > 0) {{ list[0] = 99; }} Console.WriteLine(list.Count); Console.WriteLine(list.Count == 0 ? 0 : list[0]);"
        );
        let expected_head = if values.is_empty() { "0" } else { "99" };
        assert_eq!(
            run_csharp(&src),
            vec![values.len().to_string(), expected_head.to_string()]
        );
    }
}

#[test]
fn array_matrix_sorted_min_max() {
    let cases: Vec<Vec<i32>> = vec![vec![9, 4, 7, 1], vec![12, 3, 3, 8], vec![5], vec![]];

    for values in cases {
        let literal = values
            .iter()
            .map(i32::to_string)
            .collect::<Vec<_>>()
            .join(", ");
        let src = format!(
            "var arr = new int[] {{ {literal} }}; System.Array.Sort(arr); Console.WriteLine(arr.Length); Console.WriteLine(arr.Length == 0 ? 0 : arr[0]); Console.WriteLine(arr.Length == 0 ? 0 : arr[arr.Length - 1]);"
        );
        let expected_first = values.iter().min().copied().unwrap_or(0);
        let expected_last = values.iter().max().copied().unwrap_or(0);
        assert_eq!(
            run_csharp(&src),
            vec![
                values.len().to_string(),
                expected_first.to_string(),
                expected_last.to_string(),
            ]
        );
    }
}

#[test]
fn dictionary_matrix_lookup_with_tryget() {
    let cases: Vec<(Vec<(&str, i32)>, &str)> = vec![
        (vec![("a", 10), ("b", 20)], "a"),
        (vec![("x", -1), ("y", 0), ("z", 42)], "z"),
        (vec![("only", 100)], "missing"),
    ];

    for (entries, query) in cases {
        let add_lines = entries
            .iter()
            .map(|(key, value)| format!(r#"dict.Add("{key}", {value});"#))
            .collect::<Vec<_>>()
            .join(" ");
        let src = format!(
            r#"using System.Collections.Generic; var dict = new Dictionary<string, int>(); {add_lines} bool found = dict.TryGetValue("{query}", out int value); Console.WriteLine(found); Console.WriteLine(found ? value : -1);"#
        );
        let found = entries.iter().any(|(key, _)| *key == query);
        let value = entries
            .iter()
            .find(|(key, _)| *key == query)
            .map(|(_, value)| *value)
            .unwrap_or(-1);
        assert_eq!(
            run_csharp(&src),
            vec![bool_text(found).to_string(), value.to_string()]
        );
    }
}

#[test]
fn dictionary_matrix_remove_and_check_count() {
    let entries: Vec<Vec<(&str, i32)>> = vec![
        vec![("a", 1), ("b", 2), ("c", 3)],
        vec![("x", 10), ("y", 20)],
    ];

    for set in entries {
        let add_lines = set
            .iter()
            .map(|(key, value)| format!(r#"dict.Add("{key}", {value});"#))
            .collect::<Vec<_>>()
            .join(" ");
        let remove_key = set[0].0;
        let src = format!(
            r#"using System.Collections.Generic; var dict = new Dictionary<string, int>(); {add_lines} dict.Remove("{remove_key}"); Console.WriteLine(dict.Count); Console.WriteLine(dict.ContainsKey("{remove_key}"));"#
        );
        assert_eq!(
            run_csharp(&src),
            vec![(set.len() - 1).to_string(), bool_text(false).to_string()]
        );
    }
}

#[test]
fn queue_matrix_fifo_dequeue_and_peek() {
    let cases: Vec<Vec<i32>> = vec![vec![1, 2, 3], vec![9, 7], vec![5]];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("queue.Enqueue({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let src = format!(
            "using System.Collections.Generic; var queue = new Queue<int>(); {add_lines} int first = queue.Dequeue(); int next = queue.Count == 0 ? 0 : queue.Peek(); Console.WriteLine(queue.Count + 1); Console.WriteLine(first); Console.WriteLine(next);"
        );
        let expected_head = values[0];
        let expected_next = if values.len() > 1 { values[1] } else { 0 };
        assert_eq!(
            run_csharp(&src),
            vec![
                values.len().to_string(),
                expected_head.to_string(),
                expected_next.to_string(),
            ]
        );
    }
}

#[test]
fn stack_matrix_lifo_pop_and_next_top() {
    let cases: Vec<Vec<i32>> = vec![vec![1, 2, 3], vec![9, 8], vec![10]];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("stack.Push({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let src = format!(
            "using System.Collections.Generic; var stack = new Stack<int>(); {add_lines} int top = stack.Pop(); int next = stack.Count == 0 ? 0 : stack.Peek(); Console.WriteLine(stack.Count + 1); Console.WriteLine(top); Console.WriteLine(next);"
        );
        let expected_top = values[values.len() - 1];
        let expected_next = if values.len() > 1 {
            values[values.len() - 2]
        } else {
            0
        };
        assert_eq!(
            run_csharp(&src),
            vec![
                values.len().to_string(),
                expected_top.to_string(),
                expected_next.to_string(),
            ]
        );
    }
}

#[test]
fn hashset_matrix_uniqueness_and_remove() {
    let cases: Vec<Vec<i32>> = vec![
        vec![1, 1, 2, 3, 3, 5],
        vec![4, 4, 4],
        vec![8, 9, 10],
        vec![],
    ];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("set.Add({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let unique_count = values.iter().cloned().collect::<HashSet<_>>().len();
        let removed = values.first().copied().unwrap_or(0);
        let src = format!(
            "using System.Collections.Generic; var set = new HashSet<int>(); {add_lines} bool had = set.Contains({removed}); set.Remove({removed}); Console.WriteLine(set.Count); Console.WriteLine(had); Console.WriteLine(set.Contains({removed}));"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                unique_count.to_string(),
                bool_text(!values.is_empty()).to_string(),
                bool_text(false).to_string(),
            ]
        );
    }
}

#[test]
fn list_matrix_foreach_projection_counts() {
    let cases: Vec<Vec<i32>> = vec![vec![1, 2, 3], vec![2, 4, 6, 8], vec![-1, 0, 2], vec![]];

    for values in cases {
        let add_lines = values
            .iter()
            .map(|value| format!("list.Add({value});"))
            .collect::<Vec<_>>()
            .join(" ");
        let expected_odd_count = values.iter().filter(|value| **value % 2 != 0).count();
        let expected_even_half_sum: i32 = values
            .iter()
            .filter(|value| **value % 2 == 0)
            .map(|value| value / 2)
            .sum();
        let src = format!(
            "using System.Collections.Generic; var list = new List<int>(); {add_lines} int oddCount = 0; int evenHalfSum = 0; foreach (int value in list) {{ if (value % 2 == 0) evenHalfSum += value / 2; else oddCount += 1; }} Console.WriteLine(oddCount); Console.WriteLine(evenHalfSum);"
        );
        assert_eq!(
            run_csharp(&src),
            vec![
                expected_odd_count.to_string(),
                expected_even_half_sum.to_string()
            ]
        );
    }
}
