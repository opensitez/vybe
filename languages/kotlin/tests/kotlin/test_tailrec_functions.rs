kotlin_run_test!(
    test_tailrec_factorial_accumulator,
    r#"
        tailrec fun factorial(n: Int, acc: Int = 1): Int {
            return if (n <= 1) acc else factorial(n - 1, n * acc)
        }

        fun main() {
            println(factorial(5))
        }
    "#,
    &["120"]
);

kotlin_run_test!(
    test_tailrec_sum_range,
    r#"
        tailrec fun sumRange(n: Int, acc: Int = 0): Int {
            return if (n <= 0) acc else sumRange(n - 1, acc + n)
        }

        fun main() {
            println(sumRange(4))
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_tailrec_gcd,
    r#"
        tailrec fun gcd(a: Int, b: Int): Int {
            return if (b == 0) a else gcd(b, a % b)
        }

        fun main() {
            println(gcd(48, 18))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_tailrec_string_reverse_length_based,
    r#"
        tailrec fun reverseDistance(value: String, idx: Int = 0): Int {
            return if (idx == value.length) 0 else 1 + reverseDistance(value, idx + 1)
        }

        fun main() {
            println(reverseDistance("kotlin"))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_tailrec_even_or_odd_classifier,
    r#"
        tailrec fun parity(n: Int): String {
            return if (n == 0) "even" else if (n == 1) "odd" else parity(n - 2)
        }

        fun main() {
            println(parity(9))
        }
    "#,
    &["odd"]
);

kotlin_run_test!(
    test_tailrec_list_count,
    r#"
        tailrec fun itemCount(items: List<Int>, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= items.size) acc else itemCount(items, idx + 1, acc + 1)
        }

        fun main() {
            println(itemCount(listOf(1, 2, 3, 4)))
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_tailrec_find_first_non_zero,
    r#"
        tailrec fun firstNonZero(values: List<Int>, idx: Int = 0): Int {
            return if (idx >= values.size) -1 else if (values[idx] != 0) values[idx] else firstNonZero(values, idx + 1)
        }

        fun main() {
            println(firstNonZero(listOf(0, 0, 9, 1)))
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_tailrec_power_with_step,
    r#"
        tailrec fun power(base: Int, exp: Int, acc: Int = 1): Int {
            return if (exp == 0) acc else power(base, exp - 1, acc * base)
        }

        fun main() {
            println(power(3, 4))
        }
    "#,
    &["81"]
);

kotlin_run_test!(
    test_tailrec_string_match,
    r#"
        tailrec fun countChars(text: String, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= text.length) acc else countChars(text, idx + 1, acc + if (text[idx] == 'a') 1 else 0)
        }

        fun main() {
            println(countChars("abracadabra"))
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_tailrec_binary_length,
    r#"
        tailrec fun binaryLen(values: IntArray, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= values.size) acc else binaryLen(values, idx + 1, acc + values[idx])
        }

        fun main() {
            println(binaryLen(intArrayOf(1, 2, 3, 4)))
        }
    "#,
    &["10"]
);
