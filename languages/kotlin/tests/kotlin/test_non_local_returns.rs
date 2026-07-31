kotlin_run_test!(
    test_non_local_return_from_for_each,
    r#"
        fun firstPositive(values: List<Int>): Int {
            values.forEach {
                if (it > 0) return it
            }
            return -1
        }

        fun main() {
            println(firstPositive(listOf(-2, 0, 3, 4)) )
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_local_return_from_lambda_with_label,
    r#"
        fun sumEven(values: List<Int>): Int {
            var total = 0
            values.forEach {
                if (it % 2 == 0) return@forEach
                total += it
            }
            return total
        }

        fun main() {
            println(sumEven(listOf(1, 2, 3, 4, 5)))
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_anonymous_function_return_is_local,
    r#"
        fun total(values: List<Int>): Int {
            var total = 0
            values.forEach(fun(value: Int) {
                if (value < 0) {
                    return
                }
                total += value
            })
            return total
        }

        fun main() {
            println(total(listOf(-1, 2, -3, 4)))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_nested_non_local_return_with_for_loop,
    r#"
        fun firstOdd(values: List<Int>): Int {
            values.forEach {
                run {
                    if (it % 2 == 1) {
                        return it
                    }
                }
            }
            return -1
        }

        fun main() {
            println(firstOdd(listOf(2, 4, 6, 9, 10)))
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_return_inside_map_transform,
    r#"
        fun total(values: List<Int>): Int {
            values.map {
                if (it > 10) return 100
                it
            }
            return -1
        }

        fun main() {
            println(total(listOf(3, 8, 12)))
        }
    "#,
    &["100"]
);

kotlin_run_test!(
    test_multiple_returns_from_nested_lambdas,
    r#"
        fun firstDivisible(values: List<Int>): Int {
            values.forEach {
                if (it % 3 == 0) {
                    return it
                }
            }
            return -1
        }

        fun all(values: List<Int>): Int {
            values.forEach { first ->
                if (first == 0) return 0
                if (first > 0) return@forEach
            }
            return 9
        }

        fun main() {
            println(firstDivisible(listOf(2, 5, 6, 7)))
            println(all(listOf(-1, 1, 2)))
        }
    "#,
    &["6", "9"]
);

kotlin_run_test!(
    test_non_local_return_from_fold_like_manual,
    r#"
        fun firstLarge(values: List<Int>): Int {
            values.forEach {
                if (it > 10) {
                    return it
                }
            }
            return -10
        }

        fun main() {
            println(firstLarge(listOf(3, 11, 12)))
            println(firstLarge(listOf(1, 2, 3)))
        }
    "#,
    &["11", "-10"]
);

kotlin_run_test!(
    test_return_in_for_each_indexed_non_local,
    r#"
        fun findOdd(values: IntArray): Int {
            values.forEachIndexed { index, value ->
                if (index == 2) return value
            }
            return -1
        }

        fun main() {
            println(findOdd(intArrayOf(2, 4, 6, 7, 8)))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_return_label_preserves_outer_loop,
    r#"
        fun main() {
            var count = 0
            outer@ for (i in 1..4) {
                listOf(1, 2).forEach {
                    if (i == 3 && it == 1) return@outer
                    count += i
                }
            }
            println(count)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_non_local_return_from_take_if,
    r#"
        fun firstMatching(values: List<Int>): Int {
            values.takeIf { it.isNotEmpty() }?.forEach {
                if (it == 2) return it
            }
            return -1
        }

        fun main() {
            println(firstMatching(listOf(1, 2, 3)))
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_non_local_return_with_early_false,
    r#"
        fun firstPositive(values: List<Int>): Int {
            values.forEach {
                if (it < 0) {
                    return 0
                }
            }
            return 1
        }

        fun main() {
            println(firstPositive(listOf(1, 2, -3, 4)))
            println(firstPositive(listOf(1, 2, 3)))
        }
    "#,
    &["0", "1"]
);
