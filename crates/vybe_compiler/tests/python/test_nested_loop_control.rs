use crate::helpers::run_python_one;

#[test]
fn break_inner_for_leaves_outer_running() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(2):\n for j in range(3):\n  if j == 1:\n   break\n  out.append(j)\nprint(out)\n"
        ),
        "[0, 0]"
    );
}

#[test]
fn continue_inner_for_skips_value() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(2):\n for j in range(3):\n  if j == 1:\n   continue\n  out.append(j)\nprint(out)\n"
        ),
        "[0, 2, 0, 2]"
    );
}

#[test]
fn break_outer_via_flag_pattern() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(3):\n for j in range(3):\n  if i == 1 and j == 1:\n   out.append('stop')\n   break\n  out.append(i)\n else:\n  continue\n break\nprint(out)\n"
        ),
        "[0, 0, 0, 1, 'stop']"
    );
}

#[test]
fn for_else_runs_when_inner_never_breaks() {
    assert_eq!(
        run_python_one("for i in range(2):\n for j in range(2):\n  pass\nelse:\n print('ok')\n"),
        "ok"
    );
}

#[test]
fn for_else_skipped_when_inner_breaks() {
    assert_eq!(
        run_python_one(
            "for i in range(2):\n for j in range(2):\n  if j == 1:\n   break\nelse:\n print('no')\n"
        ),
        "no"
    );
}

#[test]
fn while_nested_break_on_condition() {
    assert_eq!(
        run_python_one(
            "i = 0\nout = []\nwhile i < 3:\n j = 0\n while j < 3:\n  if j == 2:\n   break\n  out.append(i * 10 + j)\n  j += 1\n i += 1\nprint(out)\n"
        ),
        "[0, 1, 10, 11, 20, 21]"
    );
}

#[test]
fn while_nested_continue_skips_even_j() {
    assert_eq!(
        run_python_one(
            "i = 0\nout = []\nwhile i < 2:\n j = 0\n while j < 4:\n  j += 1\n  if j % 2 == 0:\n   continue\n  out.append(i * 10 + j)\n i += 1\nprint(out)\n"
        ),
        "[1, 3, 11, 13]"
    );
}

#[test]
fn triple_nested_accumulate_until_limit() {
    assert_eq!(
        run_python_one(
            "n = 0\nfor a in range(2):\n for b in range(2):\n  for c in range(2):\n   n += 1\nprint(n)\n"
        ),
        "8"
    );
}

#[test]
fn break_in_inner_while_inside_for() {
    assert_eq!(
        run_python_one(
            "out = []\nfor x in [1, 2]:\n y = 0\n while y < 5:\n  if y == 2:\n   break\n  out.append(x * 10 + y)\n  y += 1\nprint(out)\n"
        ),
        "[10, 11, 20, 21]"
    );
}

#[test]
fn continue_in_inner_while_inside_for() {
    assert_eq!(
        run_python_one(
            "out = []\nfor x in [1, 2]:\n y = 0\n while y < 4:\n  y += 1\n  if y == 2:\n   continue\n  out.append(x + y)\nprint(out)\n"
        ),
        "[2, 4, 5, 3, 5, 6]"
    );
}

#[test]
fn nested_loop_with_else_on_inner_only() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(2):\n for j in range(2):\n  out.append(i + j)\n else:\n  out.append(9)\nprint(out)\n"
        ),
        "[0, 1, 9, 1, 2, 9]"
    );
}

#[test]
fn break_first_match_search_matrix() {
    assert_eq!(
        run_python_one(
            "grid = [[0, 0], [0, 5], [0, 0]]\nfound = None\nfor r in range(3):\n for c in range(2):\n  if grid[r][c] == 5:\n   found = (r, c)\n   break\n if found:\n  break\nprint(found)\n"
        ),
        "(1, 1)"
    );
}

#[test]
fn nested_enumerate_with_break() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i, row in enumerate([[1, 2], [3, 4]]):\n for j, v in enumerate(row):\n  if v == 3:\n   out.append(i * 10 + j)\n   break\nprint(out)\n"
        ),
        "[10]"
    );
}

#[test]
fn nested_zip_with_early_break() {
    assert_eq!(
        run_python_one(
            "out = []\nfor a, b in zip([1, 2, 3], ['x', 'y', 'z']):\n for c in range(2):\n  if c == 1:\n   break\n  out.append(str(a) + b + str(c))\nprint(out)\n"
        ),
        "['1x0', '2y0', '3z0']"
    );
}

#[test]
fn while_else_after_inner_break_prevents_outer_else() {
    assert_eq!(
        run_python_one(
            "n = 0\nwhile n < 2:\n m = 0\n while m < 2:\n  if m == 1:\n   break\n  m += 1\n else:\n  print('inner')\n  n += 1\n  continue\n break\nelse:\n print('outer')\n"
        ),
        ""
    );
}

#[test]
fn nested_loop_counts_pairs_skip_diagonal() {
    assert_eq!(
        run_python_one(
            "n = 0\nfor i in range(3):\n for j in range(3):\n  if i == j:\n   continue\n  n += 1\nprint(n)\n"
        ),
        "6"
    );
}

#[test]
fn break_from_inner_for_on_string_chars() {
    assert_eq!(
        run_python_one(
            "out = []\nfor ch in 'abc':\n for _ in range(3):\n  if ch == 'b':\n   break\n  out.append(ch)\nprint(out)\n"
        ),
        "['a', 'a', 'a', 'c', 'c', 'c']"
    );
}

#[test]
fn nested_range_product_with_break_on_target() {
    assert_eq!(
        run_python_one(
            "target = 5\nfound = False\nfor a in range(3):\n for b in range(3):\n  if a * 3 + b == target:\n   found = True\n   break\n if found:\n  break\nprint(found)\n"
        ),
        "True"
    );
}

#[test]
fn continue_outer_simulated_with_flag() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(4):\n skip = False\n for j in range(2):\n  if i == 2:\n   skip = True\n   break\n if skip:\n  continue\n out.append(i)\nprint(out)\n"
        ),
        "[0, 1, 3]"
    );
}

#[test]
fn nested_while_find_first_power_of_two_above() {
    assert_eq!(
        run_python_one(
            "v = 1\nwhile v < 20:\n p = 1\n while p < v:\n  p *= 2\n if p == v:\n  print(v)\n  break\n v += 1\nelse:\n print('none')\n"
        ),
        "1"
    );
}

#[test]
fn break_in_list_iteration_nested_in_dict_keys() {
    assert_eq!(
        run_python_one(
            "d = {'a': [1, 2], 'b': [3]}\nout = []\nfor k in d:\n for v in d[k]:\n  if v == 2:\n   break\n  out.append(k + str(v))\nprint(out)\n"
        ),
        "['a1', 'b3']"
    );
}

#[test]
fn nested_loop_break_does_not_skip_outer_increment() {
    assert_eq!(
        run_python_one(
            "total = 0\nfor i in range(3):\n for j in range(5):\n  if j == 2:\n   break\n  total += 1\nprint(total)\n"
        ),
        "6"
    );
}

#[test]
fn nested_continue_only_affects_inner_index() {
    assert_eq!(
        run_python_one(
            "pairs = []\nfor a in range(2):\n for b in range(4):\n  if b % 2 == 1:\n   continue\n  pairs.append((a, b))\nprint(pairs)\n"
        ),
        "[(0, 0), (0, 2), (1, 0), (1, 2)]"
    );
}

#[test]
fn while_true_inner_break_counter() {
    assert_eq!(
        run_python_one(
            "n = 0\nwhile True:\n while True:\n  n += 1\n  if n == 3:\n   break\n break\nprint(n)\n"
        ),
        "3"
    );
}

#[test]
fn nested_for_else_break_on_last_outer_iteration() {
    assert_eq!(
        run_python_one(
            "out = []\nfor i in range(3):\n for j in range(2):\n  if i == 2 and j == 1:\n   break\n  out.append(i)\n else:\n  out.append(99)\nprint(out)\n"
        ),
        "[0, 0, 99, 1, 1, 99, 2]"
    );
}
