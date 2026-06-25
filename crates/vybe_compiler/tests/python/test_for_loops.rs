use crate::helpers::{run_python_one, run_python};

#[test]
fn for_range_counts_zero_through_four() {
    let out = run_python("for i in range(5):\n    print(i)\n");
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

#[test]
fn for_range_with_explicit_start_index() {
    let out = run_python("for i in range(2, 5):\n    print(i)\n");
    assert_eq!(out, vec!["2", "3", "4"]);
}

#[test]
fn for_range_with_step_of_two() {
    let out = run_python("for i in range(0, 10, 2):\n    print(i)\n");
    assert_eq!(out, vec!["0", "2", "4", "6", "8"]);
}

#[test]
fn for_range_with_step_of_three() {
    let out = run_python("for i in range(0, 10, 3):\n    print(i)\n");
    assert_eq!(out, vec!["0", "3", "6", "9"]);
}

#[test]
fn for_range_empty_when_start_equals_stop() {
    let out = run_python(
        "for i in range(5, 5):\n    print(i)\nprint('done')\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn for_range_single_iteration() {
    let out = run_python("for i in range(1):\n    print(i)\n");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn for_range_accumulates_running_total() {
    let out = run_python(
        "total = 0\nfor i in range(1, 5):\n    total += i\nprint(total)\n",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn for_list_iterates_all_elements() {
    let out = run_python("for x in [10, 20, 30]:\n    print(x)\n");
    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn for_list_empty_iterable_prints_nothing() {
    let out = run_python("for x in []:\n    print(x)\nprint('end')\n");
    assert_eq!(out, vec!["end"]);
}

#[test]
fn for_list_break_stops_on_target_value() {
    let out = run_python(
        "for x in [1, 2, 3, 4, 5]:\n    if x == 3:\n        break\n    print(x)\n",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn for_list_continue_skips_matching_value() {
    let out = run_python(
        "for x in [1, 2, 3, 4]:\n    if x == 2:\n        continue\n    print(x)\n",
    );
    assert_eq!(out, vec!["1", "3", "4"]);
}

#[test]
fn for_string_iterates_characters() {
    let out = run_python("for ch in 'abc':\n    print(ch)\n");
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn for_string_empty_skips_body() {
    let out = run_python("for ch in '':\n    print(ch)\nprint('ok')\n");
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn for_string_builds_word_via_concatenation() {
    let out = run_python(
        "word = ''\nfor ch in 'hi':\n    word += ch\nprint(word)\n",
    );
    assert_eq!(out, vec!["hi"]);
}

#[test]
fn for_dict_keys_iterate_single_entry() {
    let out = run_python("for key in {'only': 7}:\n    print(key)\n");
    assert_eq!(out, vec!["only"]);
}

#[test]
fn for_dict_items_unpack_key_and_value() {
    let out = run_python(
        "for key, value in {'a': 1, 'b': 2}.items():\n    print(key, value)\n",
    );
    assert_eq!(out, vec!["a 1", "b 2"]);
}

#[test]
fn for_dict_values_iterate_numbers() {
    let out = run_python("for value in {'x': 4, 'y': 5}.values():\n    print(value)\n");
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn for_dict_sums_all_values() {
    let out = run_python(
        "total = 0\nfor v in {'a': 1, 'b': 2, 'c': 3}.values():\n    total += v\nprint(total)\n",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn for_enumerate_yields_index_and_item() {
    let out = run_python("for i, item in enumerate(['a', 'b', 'c']):\n    print(i, item)\n");
    assert_eq!(out, vec!["0 a", "1 b", "2 c"]);
}

#[test]
fn for_enumerate_with_custom_start_index() {
    let out = run_python("for i, item in enumerate(['x', 'y'], start=10):\n    print(i, item)\n");
    assert_eq!(out, vec!["10 x", "11 y"]);
}

#[test]
fn for_enumerate_over_string_characters() {
    let out = run_python("for i, ch in enumerate('ab'):\n    print(i, ch)\n");
    assert_eq!(out, vec!["0 a", "1 b"]);
}

#[test]
fn for_break_exits_before_last_element() {
    let out = run_python(
        "for i in range(10):\n    if i == 3:\n        break\n    print(i)\n",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_break_on_first_iteration() {
    let out = run_python(
        "for i in range(5):\n    print(i)\n    break\nprint('after')\n",
    );
    assert_eq!(out, vec!["0", "after"]);
}

#[test]
fn for_continue_skips_even_numbers() {
    let out = run_python(
        "for i in range(6):\n    if i % 2 == 0:\n        continue\n    print(i)\n",
    );
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn for_continue_then_prints_remaining_values() {
    let out = run_python(
        "for i in range(5):\n    if i == 2:\n        continue\n    print(i)\n",
    );
    assert_eq!(out, vec!["0", "1", "3", "4"]);
}

#[test]
fn for_else_runs_when_loop_completes_without_break() {
    let out = run_python(
        "for x in [1, 2, 3]:\n    if x == 5:\n        break\nelse:\n    print('finished')\n",
    );
    assert_eq!(out, vec!["finished"]);
}

#[test]
fn for_else_skipped_when_break_occurs() {
    let out = run_python(
        "for x in [1, 2, 3]:\n    if x == 2:\n        break\nelse:\n    print('skipped')\nprint('done')\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn for_else_runs_on_empty_range() {
    let out = run_python(
        "for x in range(0):\n    print(x)\nelse:\n    print('empty')\n",
    );
    assert_eq!(out, vec!["empty"]);
}

#[test]
fn for_nested_loops_print_cartesian_pairs() {
    let out = run_python(
        "for i in range(2):\n    for j in range(2):\n        print(i, j)\n",
    );
    assert_eq!(out, vec!["0 0", "0 1", "1 0", "1 1"]);
}

#[test]
fn for_nested_three_levels_counts_iterations() {
    let out = run_python(
        "count = 0\nfor a in range(2):\n    for b in range(2):\n        for c in range(2):\n            count += 1\nprint(count)\n",
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn for_nested_inner_break_stops_inner_only() {
    let out = run_python(
        "for i in range(2):\n    for j in range(3):\n        if j == 1:\n            break\n        print(i, j)\n",
    );
    assert_eq!(out, vec!["0 0", "1 0"]);
}

#[test]
fn for_nested_inner_continue_skips_column() {
    let out = run_python(
        r#"
for i in range(3):
    for j in range(3):
        if j == 1:
            continue
        if i == 2:
            break
        print(i, j)
"#,
    );
    assert_eq!(out, vec!["0 0", "0 2", "1 0", "1 2"]);
}

#[test]
fn for_range_two_to_five_exclusive() {
    let out = run_python("for i in range(2, 6):\n    print(i)\n");
    assert_eq!(out, vec!["2", "3", "4", "5"]);
}

#[test]
fn for_list_counts_elements_matching_predicate() {
    let out = run_python(
        "count = 0\nfor n in [1, 2, 3, 4, 5]:\n    if n % 2 == 0:\n        count += 1\nprint(count)\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn for_string_counts_specific_letter() {
    let out = run_python(
        "count = 0\nfor ch in 'banana':\n    if ch == 'a':\n        count += 1\nprint(count)\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn for_dict_break_on_matching_key() {
    let out = run_python(
        "for key in {'a': 1, 'stop': 2, 'c': 3}:\n    if key == 'stop':\n        break\n    print(key)\n",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn for_enumerate_break_midway() {
    let out = run_python(
        "for i, val in enumerate([10, 20, 30, 40]):\n    if i == 2:\n        break\n    print(val)\n",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn for_list_else_with_break_on_missing_target() {
    let out = run_python(
        "for x in [4, 5, 6]:\n    if x == 1:\n        break\nelse:\n    print('all checked')\n",
    );
    assert_eq!(out, vec!["all checked"]);
}

#[test]
fn for_range_variable_upper_bound() {
    let out = run_python("n = 4\nfor i in range(n):\n    print(i)\n");
    assert_eq!(out, vec!["0", "1", "2", "3"]);
}

#[test]
fn for_list_nested_accumulates_products() {
    let out = run_python(
        "total = 0\nfor row in [[1, 2], [3, 4]]:\n    for x in row:\n        total += x\nprint(total)\n",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn for_string_uppercase_transform_in_loop() {
    let out = run_python(
        "result = ''\nfor ch in 'ab':\n    result += ch.upper()\nprint(result)\n",
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn for_dict_items_filter_by_value() {
    let out = run_python(
        "picked = 0\nfor key, value in {'a': 1, 'b': 10, 'c': 3}.items():\n    if value > 5:\n        picked += 1\nprint(picked)\n",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn for_range_break_before_any_output() {
    let out = run_python(
        "for i in range(5):\n    if i == 0:\n        break\n    print(i)\nprint('end')\n",
    );
    assert_eq!(out, vec!["end"]);
}

#[test]
fn for_list_continue_skips_all_multiples_of_three() {
    let out = run_python(
        "for n in [1, 2, 3, 4, 5, 6, 7, 8, 9]:\n    if n % 3 == 0:\n        continue\n    print(n)\n",
    );
    assert_eq!(out, vec!["1", "2", "4", "5", "7", "8"]);
}

#[test]
fn run_python_one_for_range_len_via_sum() {
    assert_eq!(
        run_python_one("print(sum([1 for x in range(4)]))\n"),
        "6"
    );
}
