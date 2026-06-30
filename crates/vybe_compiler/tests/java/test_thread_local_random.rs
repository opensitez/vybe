/// java.util.concurrent.ThreadLocalRandom — distinct bounded and unbounded draws.
use crate::helpers::{run_in_main, run_main};

#[test]
fn thread_local_random_current_is_not_null() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng != null);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_current_returns_same_instance_in_one_thread() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom a = java.util.concurrent.ThreadLocalRandom.current(); java.util.concurrent.ThreadLocalRandom b = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(a == b);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_with_bound_one_always_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(1));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_next_int_origin_equals_bound_minus_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(7, 8));",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn thread_local_random_next_int_zero_to_one_returns_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(0, 1));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_next_int_negative_origin_single_value_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(-3, -2));",
    );
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn thread_local_random_next_long_with_bound_one_always_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextLong(1L));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_next_long_origin_equals_bound_minus_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextLong(42L, 43L));",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn thread_local_random_next_double_origin_equals_bound() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextDouble(2.5, 2.5));",
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn thread_local_random_next_double_zero_to_zero_returns_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextDouble(0.0, 0.0));",
    );
    assert_eq!(out, vec!["0.0"]);
}

#[test]
fn thread_local_random_next_boolean_prints_true_or_false() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); boolean b = rng.nextBoolean(); System.out.println(b == true || b == false);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_float_is_non_negative() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); float f = rng.nextFloat(); System.out.println(f >= 0.0f);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_float_less_than_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); float f = rng.nextFloat(); System.out.println(f < 1.0f);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_double_unbounded_is_non_negative() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double d = rng.nextDouble(); System.out.println(d >= 0.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_double_unbounded_less_than_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double d = rng.nextDouble(); System.out.println(d < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_gaussian_is_finite() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double g = rng.nextGaussian(); System.out.println(Double.isFinite(g));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_unbounded_is_within_int_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int n = rng.nextInt(); System.out.println(n >= Integer.MIN_VALUE);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_long_unbounded_is_within_long_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long n = rng.nextLong(); System.out.println(n >= Long.MIN_VALUE);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_ints_limited_stream_has_expected_count() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long count = rng.ints(5L).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn thread_local_random_ints_bounded_stream_first_element_in_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int first = rng.ints(1L, 4, 5).findFirst().getAsInt(); System.out.println(first);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn thread_local_random_longs_limited_stream_has_expected_count() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long count = rng.longs(3L).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn thread_local_random_longs_bounded_stream_single_value_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long first = rng.longs(1L, 9L, 10L).findFirst().getAsLong(); System.out.println(first);",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn thread_local_random_doubles_limited_stream_has_expected_count() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long count = rng.doubles(4L).count(); System.out.println(count);",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn thread_local_random_doubles_bounded_stream_fixed_origin() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double first = rng.doubles(1L, 1.25, 1.25).findFirst().getAsDouble(); System.out.println(first);",
    );
    assert_eq!(out, vec!["1.25"]);
}

#[test]
fn thread_local_random_ints_stream_sum_is_deterministic_for_single_element_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int sum = rng.ints(3L, 2, 3).sum(); System.out.println(sum);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn thread_local_random_next_int_bound_two_is_zero_or_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int n = rng.nextInt(2); System.out.println(n == 0 || n == 1);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_bound_three_is_in_zero_to_two() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int n = rng.nextInt(3); System.out.println(n >= 0 && n <= 2);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_origin_bound_width_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(100, 101));",
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn thread_local_random_next_long_bound_two_is_zero_or_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long n = rng.nextLong(2L); System.out.println(n == 0L || n == 1L);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_double_bound_range_has_positive_width() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double d = rng.nextDouble(1.0, 2.0); System.out.println(d >= 1.0 && d < 2.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_ints_unbounded_first_value_within_int_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int n = rng.ints().limit(1L).findFirst().getAsInt(); System.out.println(n <= Integer.MAX_VALUE);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_longs_unbounded_first_value_within_long_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long n = rng.longs().limit(1L).findFirst().getAsLong(); System.out.println(n <= Long.MAX_VALUE);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_doubles_unbounded_first_value_less_than_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double d = rng.doubles().limit(1L).findFirst().getAsDouble(); System.out.println(d < 1.0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_child_thread_gets_different_instance_than_parent() {
    let types = r#"
        static java.util.concurrent.ThreadLocalRandom parentRng;
        static java.util.concurrent.ThreadLocalRandom childRng;
        static boolean same;
    "#;
    let out = run_in_main(
        "parentRng = java.util.concurrent.ThreadLocalRandom.current(); Thread t = new Thread(() -> { childRng = java.util.concurrent.ThreadLocalRandom.current(); same = parentRng == childRng; }); t.start(); t.join(); System.out.println(same);",
        types,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn thread_local_random_child_thread_current_is_not_null() {
    let types = r#"
        static boolean childHadRng;
    "#;
    let out = run_in_main(
        "Thread t = new Thread(() -> { childHadRng = java.util.concurrent.ThreadLocalRandom.current() != null; }); t.start(); t.join(); System.out.println(childHadRng);",
        types,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_negative_bound_single_value() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(-1));",
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn thread_local_random_next_int_large_origin_single_value_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(999999, 1000000));",
    );
    assert_eq!(out, vec!["999999"]);
}

#[test]
fn thread_local_random_next_long_negative_origin_single_value() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextLong(-10L, -9L));",
    );
    assert_eq!(out, vec!["-10"]);
}

#[test]
fn thread_local_random_next_double_negative_origin_single_value() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextDouble(-4.5, -4.5));",
    );
    assert_eq!(out, vec!["-4.5"]);
}

#[test]
fn thread_local_random_ints_bounded_stream_all_match_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); boolean ok = rng.ints(10L, 3, 4).allMatch(n -> n == 3); System.out.println(ok);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_longs_bounded_stream_all_match_range() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); boolean ok = rng.longs(6L, 11L, 12L).allMatch(n -> n == 11L); System.out.println(ok);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_doubles_bounded_stream_all_match_fixed_point() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); boolean ok = rng.doubles(2L, 0.5, 0.5).allMatch(d -> d == 0.5); System.out.println(ok);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_int_called_twice_with_bound_one_both_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextInt(1) + rng.nextInt(1));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_next_long_called_twice_with_bound_one_both_zero() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.nextLong(1L) + rng.nextLong(1L));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_ints_zero_length_stream_is_empty() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.ints(0L).count());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_longs_zero_length_stream_is_empty() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.longs(0L).count());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_doubles_zero_length_stream_is_empty() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); System.out.println(rng.doubles(0L).count());",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn thread_local_random_next_int_bound_zero_throws_illegal_argument() {
    let out = run_in_main(
        "try { java.util.concurrent.ThreadLocalRandom.current().nextInt(0); System.out.println(\"ok\"); } catch (IllegalArgumentException e) { System.out.println(\"bad bound\"); }",
        "",
    );
    assert_eq!(out, vec!["bad bound"]);
}

#[test]
fn thread_local_random_next_int_origin_not_less_than_bound_throws() {
    let out = run_in_main(
        "try { java.util.concurrent.ThreadLocalRandom.current().nextInt(5, 5); System.out.println(\"ok\"); } catch (IllegalArgumentException e) { System.out.println(\"bad range\"); }",
        "",
    );
    assert_eq!(out, vec!["bad range"]);
}

#[test]
fn thread_local_random_next_long_bound_zero_throws_illegal_argument() {
    let out = run_in_main(
        "try { java.util.concurrent.ThreadLocalRandom.current().nextLong(0L); System.out.println(\"ok\"); } catch (IllegalArgumentException e) { System.out.println(\"bad bound\"); }",
        "",
    );
    assert_eq!(out, vec!["bad bound"]);
}

#[test]
fn thread_local_random_next_double_origin_not_less_than_bound_throws() {
    let out = run_in_main(
        "try { java.util.concurrent.ThreadLocalRandom.current().nextDouble(3.0, 2.0); System.out.println(\"ok\"); } catch (IllegalArgumentException e) { System.out.println(\"bad range\"); }",
        "",
    );
    assert_eq!(out, vec!["bad range"]);
}

#[test]
fn thread_local_random_ints_negative_size_throws() {
    let out = run_in_main(
        "try { java.util.concurrent.ThreadLocalRandom.current().ints(-1L).count(); System.out.println(\"ok\"); } catch (IllegalArgumentException e) { System.out.println(\"bad size\"); }",
        "",
    );
    assert_eq!(out, vec!["bad size"]);
}

#[test]
fn thread_local_random_next_float_twice_still_less_than_one() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); float a = rng.nextFloat(); float b = rng.nextFloat(); System.out.println(a < 1.0f && b < 1.0f);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_next_gaussian_twice_both_finite() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); double a = rng.nextGaussian(); double b = rng.nextGaussian(); System.out.println(Double.isFinite(a) && Double.isFinite(b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn thread_local_random_ints_bounded_max_of_stream_is_fixed_origin() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); int max = rng.ints(8L, 6, 7).max().getAsInt(); System.out.println(max);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn thread_local_random_longs_bounded_min_of_stream_is_fixed_origin() {
    let out = run_main(
        "java.util.concurrent.ThreadLocalRandom rng = java.util.concurrent.ThreadLocalRandom.current(); long min = rng.longs(4L, 55L, 56L).min().getAsLong(); System.out.println(min);",
    );
    assert_eq!(out, vec!["55"]);
}
