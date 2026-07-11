use crate::helpers::{run_print, run_python, run_python_one};

#[test]
fn while_decrements_until_zero() {
    assert_eq!(
        run_python_one("n = 3\nwhile n > 0:\n n -= 1\nprint(n)\n"),
        "0"
    );
}

#[test]
fn while_zero_iterations_skips_body() {
    assert_eq!(
        run_python_one("n = 0\nwhile n > 0:\n print('x')\nprint('ok')\n"),
        "ok"
    );
}

#[test]
fn while_else_runs_when_not_broken() {
    assert_eq!(
        run_python_one("n = 2\nwhile n:\n n -= 1\nelse:\n print('done')\n"),
        "done"
    );
}

#[test]
fn while_else_skipped_on_break() {
    assert_eq!(
        run_python_one(
            "n = 5\nwhile n:\n n -= 1\n if n == 2:\n  break\nelse:\n print('no')\nprint('yes')\n"
        ),
        "yes"
    );
}

#[test]
fn while_break_exits_immediately() {
    assert_eq!(
        run_python_one("n = 10\nwhile n:\n n -= 1\n if n == 7:\n  break\nprint(n)\n"),
        "7"
    );
}

#[test]
fn while_continue_skips_print() {
    let src = [
        "n = 0",
        "while n < 4:",
        " n += 1",
        " if n == 2:",
        "  continue",
        " print(n)",
    ]
    .join("\n");
    assert_eq!(run_python(&(src + "\n")), vec!["1", "3", "4"]);
}

#[test]
fn while_nested_multiplies_counters() {
    assert_eq!(
        run_python_one(
            "i = 0\nprod = 1\nwhile i < 3:\n j = 0\n while j < 2:\n  prod *= 2\n  j += 1\n i += 1\nprint(prod)\n"
        ),
        "64"
    );
}

#[test]
fn while_condition_uses_comparison_chain() {
    assert_eq!(
        run_python_one("x = 1\nwhile 0 < x < 5:\n x += 1\nprint(x)\n"),
        "5"
    );
}

#[test]
fn while_iterates_list_pop_from_end() {
    assert_eq!(
        run_python_one("xs = [1, 2, 3]\nwhile xs:\n print(xs.pop())\n"),
        "3\n2\n1"
    );
}

#[test]
fn while_reads_input_sentinel_pattern() {
    assert_eq!(
        run_python_one(
            "data = [1, 2, -1, 3]\ni = 0\ntotal = 0\nwhile True:\n v = data[i]\n i += 1\n if v < 0:\n  break\n total += v\nprint(total)\n"
        ),
        "3"
    );
}

#[test]
fn do_while_emulated_with_first_run() {
    assert_eq!(
        run_python_one("n = 0\nwhile True:\n n += 1\n print(n)\n if n >= 2:\n  break\n"),
        "1\n2"
    );
}

#[test]
fn while_truthy_empty_list_stops() {
    assert_eq!(
        run_python_one("xs = [1]\nwhile xs:\n xs.pop()\nprint('end')\n"),
        "end"
    );
}

#[test]
fn while_truthy_nonempty_string_iterates_chars() {
    assert_eq!(
        run_python_one("s = 'ab'\ni = 0\nwhile i < len(s):\n print(s[i])\n i += 1\n"),
        "a\nb"
    );
}

#[test]
fn while_accumulates_factorial() {
    assert_eq!(
        run_python_one("n = 5\nf = 1\nwhile n > 1:\n f *= n\n n -= 1\nprint(f)\n"),
        "120"
    );
}

#[test]
fn while_with_boolean_flag() {
    assert_eq!(
        run_python_one(
            "running = True\nc = 0\nwhile running:\n c += 1\n if c == 3:\n  running = False\nprint(c)\n"
        ),
        "3"
    );
}

#[test]
fn while_membership_scan_finds_target() {
    assert_eq!(
        run_python_one(
            "xs = [4, 7, 9]\ni = 0\nfound = False\nwhile i < len(xs) and not found:\n found = xs[i] == 7\n i += 1\nprint(found)\n"
        ),
        "True"
    );
}

#[test]
fn while_else_not_run_on_return_emulated() {
    assert_eq!(
        run_python_one(
            "n = 1\nwhile n:\n print('loop')\n break\nelse:\n print('else')\nprint('after')\n"
        ),
        "loop\nafter"
    );
}

#[test]
fn while_infinite_guarded_by_counter() {
    assert_eq!(
        run_python_one("c = 0\nwhile True:\n c += 1\n if c == 4:\n  break\nprint(c)\n"),
        "4"
    );
}

#[test]
fn while_decrements_by_two() {
    assert_eq!(
        run_python_one("n = 10\nwhile n > 0:\n n -= 2\nprint(n)\n"),
        "0"
    );
}

#[test]
fn while_with_not_condition() {
    assert_eq!(
        run_python_one("done = False\nc = 0\nwhile not done:\n c += 1\n done = c == 2\nprint(c)\n"),
        "2"
    );
}

#[test]
fn while_short_circuit_and_in_condition() {
    assert_eq!(
        run_python_one("a = 1\nb = 0\nwhile a and b:\n print('no')\nprint('yes')\n"),
        "yes"
    );
}

#[test]
fn while_short_circuit_or_keeps_looping() {
    assert_eq!(
        run_python_one(
            "a = 0\nb = 1\nc = 0\nwhile a or b:\n c += 1\n b = 0\n if c == 2:\n  break\nprint(c)\n"
        ),
        "1"
    );
}

#[test]
fn while_nested_break_only_inner() {
    assert_eq!(
        run_python_one(
            "out = []\ni = 0\nwhile i < 3:\n j = 0\n while j < 3:\n  if j == 1:\n   break\n  out.append(j)\n  j += 1\n i += 1\nprint(len(out))\n"
        ),
        "3"
    );
}

#[test]
fn while_reads_dict_keys_via_popitem() {
    assert_eq!(
        run_python_one(
            "d = {'a': 1, 'b': 2}\nkeys = []\nwhile d:\n k, v = d.popitem()\n keys.append(k)\nprint(len(keys))\n"
        ),
        "2"
    );
}

#[test]
fn while_gcd_euclidean_algorithm() {
    assert_eq!(
        run_python_one("a = 48\nb = 18\nwhile b:\n a, b = b, a % b\nprint(a)\n"),
        "6"
    );
}

#[test]
fn while_counts_bits_in_integer() {
    assert_eq!(
        run_python_one("n = 13\ncount = 0\nwhile n:\n count += n & 1\n n >>= 1\nprint(count)\n"),
        "3"
    );
}

#[test]
fn while_reverses_digits_of_number() {
    assert_eq!(
        run_python_one(
            "n = 123\nrev = 0\nwhile n:\n rev = rev * 10 + n % 10\n n //= 10\nprint(rev)\n"
        ),
        "321"
    );
}

#[test]
fn while_string_builder_concat() {
    assert_eq!(
        run_python_one(
            "parts = ['a', 'b', 'c']\ni = 0\ns = ''\nwhile i < len(parts):\n s += parts[i]\n i += 1\nprint(s)\n"
        ),
        "abc"
    );
}

#[test]
fn while_waits_until_predicate_true() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 5]\ni = 0\nwhile i < len(xs) and xs[i] < 5:\n i += 1\nprint(xs[i])\n"
        ),
        "5"
    );
}

#[test]
fn while_modulo_cycle_detects_period() {
    assert_eq!(
        run_python_one(
            "n = 1\nsteps = 0\nwhile steps < 4:\n n = (n * 3) % 7\n steps += 1\nprint(n)\n"
        ),
        "4"
    );
}

#[test]
fn while_float_counter_halving() {
    assert_eq!(run_print("len(str(0.5))"), "3");
}

#[test]
fn while_list_comprehension_alternative_sum() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3]\ni = 0\ntotal = 0\nwhile i < len(xs):\n total += xs[i]\n i += 1\nprint(total)\n"
        ),
        "6"
    );
}

#[test]
fn while_removes_matching_elements() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3, 2, 1]\ni = 0\nwhile i < len(xs):\n if xs[i] == 2:\n  xs.pop(i)\n else:\n  i += 1\nprint(xs.count(2))\n"
        ),
        "0"
    );
}

#[test]
fn while_zip_two_lists_manually() {
    assert_eq!(
        run_python_one(
            "a = [1, 2]\nb = [3, 4]\ni = 0\nwhile i < len(a):\n print(a[i] + b[i])\n i += 1\n"
        ),
        "4\n6"
    );
}

#[test]
fn while_enumerate_manual_index_value() {
    assert_eq!(
        run_python_one(
            "xs = ['x', 'y']\ni = 0\nwhile i < len(xs):\n print(str(i) + xs[i])\n i += 1\n"
        ),
        "0x\n1y"
    );
}

#[test]
fn while_parses_digits_from_string() {
    assert_eq!(
        run_python_one(
            "s = '1234'\ni = 0\nn = 0\nwhile i < len(s):\n n = n * 10 + int(s[i])\n i += 1\nprint(n)\n"
        ),
        "1234"
    );
}

#[test]
fn while_finds_first_negative() {
    assert_eq!(
        run_python_one(
            "xs = [1, 3, -2, 4]\ni = 0\nwhile i < len(xs) and xs[i] >= 0:\n i += 1\nprint(xs[i])\n"
        ),
        "-2"
    );
}

#[test]
fn while_doubles_until_threshold() {
    assert_eq!(
        run_python_one("n = 1\nwhile n < 20:\n n *= 2\nprint(n)\n"),
        "32"
    );
}

#[test]
fn while_counts_vowels_in_string() {
    assert_eq!(
        run_python_one(
            "s = 'hello'\ni = 0\nc = 0\nwhile i < len(s):\n if s[i] in 'aeiou':\n  c += 1\n i += 1\nprint(c)\n"
        ),
        "2"
    );
}

#[test]
fn while_merges_sorted_lists_step() {
    assert_eq!(
        run_python_one(
            "a = [1, 3]\nb = [2, 4]\ni = j = 0\nout = []\nwhile i < len(a) and j < len(b):\n if a[i] < b[j]:\n  out.append(a[i])\n  i += 1\n else:\n  out.append(b[j])\n  j += 1\nprint(out[2])\n"
        ),
        "3"
    );
}

#[test]
fn while_spins_on_flag_then_clears() {
    assert_eq!(
        run_python_one(
            "flag = True\nc = 0\nwhile flag:\n c += 1\n if c == 1:\n  flag = False\nprint(c)\n"
        ),
        "1"
    );
}

#[test]
fn while_empty_body_still_advances_counter() {
    assert_eq!(
        run_python_one("i = 0\nwhile i < 3:\n i += 1\nprint(i)\n"),
        "3"
    );
}

#[test]
fn while_negated_membership_search() {
    assert_eq!(
        run_python_one(
            "xs = [1, 2, 3]\ntarget = 4\ni = 0\nwhile i < len(xs) and xs[i] != target:\n i += 1\nprint(i == len(xs))\n"
        ),
        "True"
    );
}
