use crate::helpers::{run_in_main, run_main};

#[test]
fn varargs_method_sums_four_arguments() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(1, 2, 3, 4));", types);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn varargs_method_with_no_arguments_returns_zero() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum());", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn varargs_method_with_single_argument_returns_it() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(5));", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn varargs_accepts_explicit_int_array_argument() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(new int[] {2, 3, 5}));", types);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn string_varargs_joins_with_dash() {
    let types = r#"
        static String join(String... parts) {
            String out = \"\";
            for (int i = 0; i < parts.length; i++) {
                if (i > 0) out = out + \"-\";
                out = out + parts[i];
            }
            return out;
        }
    "#;
    let out = run_in_main("System.out.println(join(\"a\", \"b\", \"c\"));", types);
    assert_eq!(out, vec!["a-b-c"]);
}

#[test]
fn varargs_after_required_parameter_uses_both() {
    let types = r#"
        static int prefixSum(int head, int... tail) {
            int total = head;
            for (int n : tail) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(prefixSum(10, 1, 2));", types);
    assert_eq!(out, vec!["13"]);
}

#[test]
fn varargs_loop_counts_received_arguments() {
    let types = r#"
        static int count(int... nums) {
            int c = 0;
            for (int n : nums) c = c + 1;
            return c;
        }
    "#;
    let out = run_in_main("System.out.println(count(1, 2, 3, 4, 5));", types);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn varargs_all_zero_arguments_sum_to_zero() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(0, 0, 0));", types);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn varargs_negative_numbers_sum_correctly() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(-2, 5, -1));", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn varargs_single_element_array_spreads_values() {
    let types = r#"
        static int first(int... nums) { return nums[0]; }
    "#;
    let out = run_in_main("System.out.println(first(new int[] {42}));", types);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn string_varargs_empty_call_returns_empty_string() {
    let types = r#"
        static String concat(String... parts) {
            String out = \"\";
            for (String p : parts) out = out + p;
            return out;
        }
    "#;
    let out = run_in_main("System.out.println(concat());", types);
    assert_eq!(out, vec![""]);
}

#[test]
fn varargs_max_picks_largest_argument() {
    let types = r#"
        static int maxOf(int... nums) {
            int best = nums[0];
            for (int n : nums) if (n > best) best = n;
            return best;
        }
    "#;
    let out = run_in_main("System.out.println(maxOf(3, 9, 1));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn varargs_first_element_read_by_index() {
    let types = r#"
        static int first(int... nums) { return nums[0]; }
    "#;
    let out = run_in_main("System.out.println(first(7, 8, 9));", types);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn varargs_last_element_read_by_index() {
    let types = r#"
        static int last(int... nums) { return nums[nums.length - 1]; }
    "#;
    let out = run_in_main("System.out.println(last(7, 8, 9));", types);
    assert_eq!(out, vec!["9"]);
}

#[test]
fn varargs_ten_arguments_are_all_counted() {
    let types = r#"
        static int count(int... nums) {
            int c = 0;
            for (int n : nums) c = c + 1;
            return c;
        }
    "#;
    let out = run_in_main(
        "System.out.println(count(1, 2, 3, 4, 5, 6, 7, 8, 9, 10));",
        types,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn varargs_double_values_sum_as_double() {
    let types = r#"
        static double sum(double... nums) {
            double total = 0.0;
            for (double n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(sum(1.5, 2.0));", types);
    assert_eq!(out, vec!["3.5"]);
}

#[test]
fn varargs_boolean_count_tracks_true_values() {
    let types = r#"
        static int trueCount(boolean... flags) {
            int c = 0;
            for (boolean f : flags) if (f) c = c + 1;
            return c;
        }
    "#;
    let out = run_in_main("System.out.println(trueCount(true, false, true));", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn varargs_required_prefix_without_tail_uses_head_only() {
    let types = r#"
        static int prefixSum(int head, int... tail) {
            int total = head;
            for (int n : tail) total += n;
            return total;
        }
    "#;
    let out = run_in_main("System.out.println(prefixSum(4));", types);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn instance_varargs_method_sums_fields() {
    let types = r#"
        static class Adder {
            int sum(int... nums) {
                int total = 0;
                for (int n : nums) total += n;
                return total;
            }
        }
    "#;
    let out = run_in_main(
        "Adder a = new Adder(); System.out.println(a.sum(2, 3));",
        types,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn varargs_length_property_reports_argument_count() {
    let types = r#"
        static int size(int... nums) { return nums.length; }
    "#;
    let out = run_in_main("System.out.println(size(1, 2, 3));", types);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn string_varargs_reports_argument_count() {
    let types = r#"
        static int size(String... parts) { return parts.length; }
    "#;
    let out = run_in_main("System.out.println(size(\"a\", \"b\"));", types);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn varargs_forwards_to_another_varargs_method() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
        static int twice(int... nums) { return sum(nums) + sum(nums); }
    "#;
    let out = run_in_main("System.out.println(twice(1, 2));", types);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn varargs_prints_each_argument_in_order() {
    let types = r#"
        static void show(int... nums) {
            for (int n : nums) System.out.println(n);
        }
    "#;
    let out = run_in_main("show(4, 5, 6);", types);
    assert_eq!(out, vec!["4", "5", "6"]);
}

#[test]
fn varargs_with_only_array_argument_preserves_order() {
    let types = r#"
        static int first(int... nums) { return nums[0]; }
        static int second(int... nums) { return nums[1]; }
    "#;
    let out = run_in_main(
        "int[] data = new int[] {11, 22}; System.out.println(first(data)); System.out.println(second(data));",
        types,
    );
    assert_eq!(out, vec!["11", "22"]);
}

#[test]
fn varargs_mixed_direct_and_array_call_match() {
    let types = r#"
        static int sum(int... nums) {
            int total = 0;
            for (int n : nums) total += n;
            return total;
        }
    "#;
    let out = run_in_main(
        "System.out.println(sum(1, 2)); System.out.println(sum(new int[] {1, 2}));",
        types,
    );
    assert_eq!(out, vec!["3", "3"]);
}
