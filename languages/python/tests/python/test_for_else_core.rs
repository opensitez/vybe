use crate::helpers::{run_python, run_python_one};

#[test]
fn for_else_runs_when_loop_completes() {
    assert_eq!(
        run_python_one("for x in range(3):\n pass\nelse:\n print('done')\n"),
        "done"
    );
}

#[test]
fn for_else_skipped_on_break() {
    assert_eq!(
        run_python_one(
            "for x in range(5):\n if x == 2:\n  break\nelse:\n print('no')\nprint('yes')\n"
        ),
        "yes"
    );
}

#[test]
fn for_else_runs_on_empty_iterable() {
    assert_eq!(
        run_python_one("for x in []:\n print('loop')\nelse:\n print('else')\n"),
        "else"
    );
}

#[test]
fn for_else_finds_no_match_sets_flag() {
    assert_eq!(
        run_python_one(
            "found = False\nfor x in [1, 3, 5]:\n if x == 2:\n  found = True\n  break\nelse:\n print('missing')\nprint(found)\n"
        ),
        "missing\nFalse"
    );
}

#[test]
fn for_else_break_on_match_skips_else() {
    assert_eq!(
        run_python_one(
            "for x in [1, 2, 3]:\n if x == 2:\n  print('hit')\n  break\nelse:\n print('else')\n"
        ),
        "hit"
    );
}

#[test]
fn for_else_search_pattern_success() {
    assert_eq!(
        run_python_one(
            "target = 7\nfor n in [1, 3, 5, 7, 9]:\n if n == target:\n  print('found')\n  break\nelse:\n print('not found')\n"
        ),
        "found"
    );
}

#[test]
fn for_else_search_pattern_failure() {
    assert_eq!(
        run_python_one(
            "target = 4\nfor n in [1, 3, 5]:\n if n == target:\n  print('found')\n  break\nelse:\n print('not found')\n"
        ),
        "not found"
    );
}

#[test]
fn for_else_with_continue_still_runs() {
    assert_eq!(
        run_python_one("for x in range(3):\n if x == 1:\n  continue\nelse:\n print('done')\n"),
        "done"
    );
}

#[test]
fn for_else_over_string_chars() {
    assert_eq!(
        run_python_one("for ch in 'ab':\n pass\nelse:\n print('end')\n"),
        "end"
    );
}

#[test]
fn for_else_over_dict_keys() {
    assert_eq!(
        run_python_one("for k in {'a': 1, 'b': 2}:\n pass\nelse:\n print('keys')\n"),
        "keys"
    );
}

#[test]
fn for_else_over_dict_items() {
    assert_eq!(
        run_python_one(
            "total = 0\nfor k, v in {'a': 1, 'b': 2}.items():\n total += v\nelse:\n print(total)\n"
        ),
        "6"
    );
}

#[test]
fn for_else_nested_outer_runs() {
    assert_eq!(
        run_python_one(
            "for i in range(2):\n for j in range(2):\n  pass\n else:\n  print(j)\nelse:\n print('outer')\n"
        ),
        "1\n1\nouter"
    );
}

#[test]
fn for_else_inner_break_does_not_trigger_outer_else() {
    assert_eq!(
        run_python_one(
            "for i in range(2):\n for j in range(3):\n  if j == 1:\n   break\n else:\n  print('inner')\nelse:\n print('outer')\n"
        ),
        "outer"
    );
}

#[test]
fn for_else_with_range_step() {
    assert_eq!(
        run_python_one("for x in range(0, 6, 2):\n pass\nelse:\n print('even steps')\n"),
        "even steps"
    );
}

#[test]
fn for_else_prime_check_small_composite() {
    assert_eq!(
        run_python_one(
            "n = 9\nfor d in range(2, n):\n if n % d == 0:\n  print('composite')\n  break\nelse:\n print('prime')\n"
        ),
        "composite"
    );
}

#[test]
fn for_else_prime_check_small_prime() {
    assert_eq!(
        run_python_one(
            "n = 7\nfor d in range(2, n):\n if n % d == 0:\n  print('composite')\n  break\nelse:\n print('prime')\n"
        ),
        "prime"
    );
}

#[test]
fn for_else_first_duplicate_finder() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3, 2, 4]\nseen = set()\nfor x in xs:\n if x in seen:\n  print('dup', x)\n  break\n seen.add(x)\nelse:\n print('unique')\n"
        ),
        "dup 2"
    );
}

#[test]
fn for_else_all_unique_list() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3]\nseen = set()\nfor x in xs:\n if x in seen:\n  print('dup')\n  break\n seen.add(x)\nelse:\n print('unique')\n"
        ),
        "unique"
    );
}

#[test]
fn for_else_enumerate_manual_search() {
    assert_eq!(
        run_python_one(
            "words = ['cat', 'dog', 'bird']\nfor i in range(len(words)):\n if words[i] == 'dog':\n  print(i)\n  break\nelse:\n print(-1)\n"
        ),
        "1"
    );
}

#[test]
fn for_else_zip_iteration_complete() {
    assert_eq!(
        run_python_one("for a, b in zip([1, 2], [3, 4]):\n pass\nelse:\n print('zipped')\n"),
        "zipped"
    );
}

#[test]
fn for_else_list_comprehension_alternative_any() {
    assert_eq!(
        run_python_one(
            "xs = [2, 4, 6]\nfor x in xs:\n if x % 2 == 1:\n  print('odd')\n  break\nelse:\n print('all even')\n"
        ),
        "all even"
    );
}

#[test]
fn for_else_finds_odd_element() {
    assert_eq!(
        run_python_one(
            "xs = [2, 4, 5]\nfor x in xs:\n if x % 2 == 1:\n  print('odd')\n  break\nelse:\n print('all even')\n"
        ),
        "odd"
    );
}

#[test]
fn for_else_over_set_elements() {
    assert_eq!(
        run_python_one("for x in {1, 2, 3}:\n pass\nelse:\n print('set')\n"),
        "set"
    );
}

#[test]
fn for_else_over_tuple() {
    assert_eq!(
        run_python_one("for x in (1, 2):\n pass\nelse:\n print('tuple')\n"),
        "tuple"
    );
}

#[test]
fn for_else_with_early_break_on_first() {
    assert_eq!(
        run_python_one("for x in [9]:\n print(x)\n break\nelse:\n print('else')\nprint('after')\n"),
        "9\nafter"
    );
}

#[test]
fn for_else_single_iteration_no_break() {
    assert_eq!(
        run_python_one("for x in [42]:\n print(x)\nelse:\n print('else')\n"),
        "42\nelse"
    );
}

#[test]
fn for_else_validates_all_predicates() {
    assert_eq!(
        run_python_one(
            "nums = [2, 4, 8]\nfor n in nums:\n if n < 0:\n  break\nelse:\n print('ok')\n"
        ),
        "ok"
    );
}

#[test]
fn for_else_predicate_fails_triggers_break() {
    assert_eq!(
        run_python_one(
            "nums = [2, -1, 8]\nfor n in nums:\n if n < 0:\n  print('bad')\n  break\nelse:\n print('ok')\n"
        ),
        "bad"
    );
}

#[test]
fn for_else_reads_file_lines_pattern() {
    assert_eq!(
        run_python_one(
            "lines = ['ok', 'done']\nfor line in lines:\n if line == 'error':\n  print('fail')\n  break\nelse:\n print('pass')\n"
        ),
        "pass"
    );
}

#[test]
fn for_else_reads_file_lines_error_line() {
    assert_eq!(
        run_python_one(
            "lines = ['ok', 'error']\nfor line in lines:\n if line == 'error':\n  print('fail')\n  break\nelse:\n print('pass')\n"
        ),
        "fail"
    );
}

#[test]
fn for_else_generator_exhausted() {
    assert_eq!(
        run_python_one(
            "def gen():\n yield 1\n yield 2\nfor x in gen():\n pass\nelse:\n print('gen')\n"
        ),
        "gen"
    );
}

#[test]
fn for_else_with_return_in_function() {
    assert_eq!(
        run_python_one(
            "def f():\n for x in range(2):\n  if x == 5:\n   break\n else:\n  return 'done'\n return 'skip'\nprint(f())\n"
        ),
        "done"
    );
}

#[test]
fn for_else_while_emulated_via_for_range() {
    assert_eq!(
        run_python_one(
            "n = 3\nfor _ in range(100):\n n -= 1\n if n <= 0:\n  break\nelse:\n print('limit')\nprint(n)\n"
        ),
        "0"
    );
}

#[test]
fn for_else_membership_all_present() {
    assert_eq!(
        run_python_one(
            "need = 'ae'\nword = 'cat'\nfor ch in need:\n if ch not in word:\n  print('missing')\n  break\nelse:\n print('has all')\n"
        ),
        "missing"
    );
}

#[test]
fn for_else_sorted_unique_check() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3]\nfor i in range(1, len(xs)):\n if xs[i] < xs[i-1]:\n  print('unsorted')\n  break\nelse:\n print('sorted')\n"
        ),
        "sorted"
    );
}

#[test]
fn for_else_unsorted_detected() {
    assert_eq!(
        run_python_one(
            "xs = [1, 3, 2]\nfor i in range(1, len(xs)):\n if xs[i] < xs[i-1]:\n  print('unsorted')\n  break\nelse:\n print('sorted')\n"
        ),
        "unsorted"
    );
}

#[test]
fn for_else_multiple_prints_in_body() {
    assert_eq!(
        run_python("for i in range(2):\n print(i)\nelse:\n print('fin')\n"),
        vec!["0", "1", "fin"]
    );
}

#[test]
fn for_else_over_reversed_range_list() {
    assert_eq!(
        run_python_one("for x in reversed([1, 2, 3]):\n pass\nelse:\n print('rev')\n"),
        "rev"
    );
}

#[test]
fn for_else_count_until_threshold() {
    assert_eq!(
        run_python_one(
            "count = 0\nfor _ in range(10):\n count += 1\n if count == 3:\n  break\nelse:\n print('full')\nprint(count)\n"
        ),
        "3"
    );
}

#[test]
fn for_else_count_completes_full_range() {
    assert_eq!(
        run_python_one(
            "count = 0\nfor _ in range(3):\n count += 1\nelse:\n print('full')\nprint(count)\n"
        ),
        "full\n3"
    );
}

#[test]
fn for_else_with_pass_only_body() {
    assert_eq!(
        run_python_one("for _ in range(2):\n pass\nelse:\n print(1)\n"),
        "1"
    );
}

#[test]
fn for_else_nested_break_only_inner_else_skipped() {
    assert_eq!(
        run_python_one(
            "for i in range(2):\n for j in range(2):\n  if j == 0:\n   break\n else:\n  print('inner-else')\nelse:\n print('outer-else')\n"
        ),
        "outer-else"
    );
}

#[test]
fn for_else_find_substring_position() {
    assert_eq!(
        run_python_one(
            "hay = 'hello'\nneedle = 'll'\nfor i in range(len(hay) - len(needle) + 1):\n if hay[i:i+len(needle)] == needle:\n  print(i)\n  break\nelse:\n print(-1)\n"
        ),
        "2"
    );
}

#[test]
fn for_else_substring_not_found() {
    assert_eq!(
        run_python_one(
            "hay = 'hello'\nneedle = 'zz'\nfor i in range(len(hay) - len(needle) + 1):\n if hay[i:i+len(needle)] == needle:\n  print(i)\n  break\nelse:\n print(-1)\n"
        ),
        "-1"
    );
}
