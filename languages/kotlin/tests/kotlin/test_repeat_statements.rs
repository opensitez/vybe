kotlin_run_test!(
    test_repeat_builds_sum,
    r#"
        fun main() {
            var sum = 0
            repeat(5) { i ->
                sum += i
            }
            println(sum)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_repeat_zero_count,
    r#"
        fun main() {
            var out = 0
            repeat(0) {
                out += 1
            }
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_repeat_variable_count,
    r#"
        fun main() {
            val n = 4
            var out = 0
            repeat(n) { i ->
                out += i + 1
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_repeat_accumulate_chars,
    r#"
        fun main() {
            var out = ""
            repeat(3) {
                out += "x"
            }
            println(out)
        }
    "#,
    &["xxx"]
);

kotlin_run_test!(
    test_repeat_nested,
    r#"
        fun main() {
            var out = 0
            repeat(2) {
                repeat(3) {
                    out += 1
                }
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_repeat_without_index,
    r#"
        fun main() {
            var out = 0
            repeat(4) {
                out = out * 10 + 1
            }
            println(out)
        }
    "#,
    &["1111"]
);

kotlin_run_test!(
    test_repeat_while_mixed,
    r#"
        fun main() {
            var i = 0
            var out = 0
            repeat(5) {
                if (i % 2 == 0) out += i
                i += 1
            }
            println(out)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_repeat_with_function,
    r#"
        fun addOne(x: Int) = x + 1
        fun main() {
            var total = 0
            repeat(4) {
                total = addOne(total)
            }
            println(total)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_repeat_string_index,
    r#"
        fun main() {
            var out = ""
            repeat(3) { index ->
                out += index.toString()
            }
            println(out)
        }
    "#,
    &["012"]
);

kotlin_run_test!(
    test_repeat_mutable_state,
    r#"
        var acc = 0
        fun bump() {
            acc += 1
        }
        fun main() {
            repeat(7) { bump() }
            println(acc)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_repeat_after_zero_like,
    r#"
        fun main() {
            var out = 1
            repeat(0) {
                out = 9
            }
            repeat(2) {
                out += out
            }
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_repeat_in_class_init,
    r#"
        class Counter {
            var total = 0
            init {
                repeat(3) { total += 1 }
            }
        }
        fun main() {
            println(Counter().total)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_repeat_in_conditioned_method,
    r#"
        class Repeater {
            fun build(a: Int): Int {
                var sum = 0
                repeat(a) { i -> sum += i }
                return sum
            }
        }
        fun main() {
            println(Repeater().build(6))
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_repeat_local_capture,
    r#"
        fun main() {
            var label = ""
            repeat(4) { i ->
                if (i > 1) label += i.toString()
            }
            println(label)
        }
    "#,
    &["23"]
);

kotlin_run_test!(
    test_repeat_with_longs,
    r#"
        fun main() {
            var total = 0L
            repeat(4) { i ->
                total += i.toLong()
            }
            println(total)
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_repeat_large,
    r#"
        fun main() {
            var out = 0
            repeat(9) {
                out += 1
            }
            println(out)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_repeat_nested_with_break_condition,
    r#"
    fun main() {
        var out = 0
        repeat(5) { outer ->
            repeat(3) { inner ->
                if (outer == 2 && inner == 2) return@repeat
                out += 1
            }
        }
        println(out)
    }
    "#,
    &["14"]
);

kotlin_run_test!(
    test_repeat_inlined_return_is_not_used,
    r#"
        fun main() {
            var out = 0
            repeat(3) {
                out += 1
            }
            println(out)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_repeat_and_boolean,
    r#"
        fun main() {
            var out = true
            repeat(2) {
                out = out && (it < 2)
            }
            println(out)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_repeat_with_string_concat_condition,
    r#"
        fun main() {
            var out = ""
            repeat(4) { n ->
                out += if (n % 2 == 0) "E" else "O"
            }
            println(out)
        }
    "#,
    &["EOEO"]
);

kotlin_run_test!(
    test_repeat_in_tail_function,
    r#"
        fun inc(n: Int): Int {
            var out = 0
            repeat(n) { out += 1 }
            return out
        }
        fun main() {
            println(inc(5))
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_repeat_with_non_zero_start,
    r#"
        fun main() {
            var i = 3
            var out = 0
            repeat(4) {
                out += i
                i += 1
            }
            println(out)
        }
    "#,
    &["22"]
);

kotlin_run_test!(
    test_repeat_with_local_block,
    r#"
        fun main() {
            var out = ""
            repeat(3) {
                out += "x"
            }
            println(out)
        }
    "#,
    &["xxx"]
);

kotlin_run_test!(
    test_repeat_long_count,
    r#"
        fun main() {
            var out = 0L
            repeat(5) { i ->
                out += i.toLong()
            }
            println(out)
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_repeat_nested_in_condition,
    r#"
        fun main() {
            var out = 0
            repeat(4) { outer ->
                if (outer == 2) {
                    repeat(2) {
                        out += outer
                    }
                }
            }
            println(out)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_repeat_boolean_flags,
    r#"
        fun main() {
            var out = true
            repeat(3) {
                out = out && it != 1
            }
            println(out)
        }
    "#,
    &["false"]
);

kotlin_run_test!(
    test_repeat_with_mutating_var,
    r#"
        class C {
            var n = 0
            fun inc() { n += 1 }
        }
        fun main() {
            val c = C()
            repeat(3) { c.inc() }
            println(c.n)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_repeat_nested_count,
    r#"
        fun main() {
            var out = 0
            repeat(2) {
                repeat(2) {
                    repeat(2) {
                        out += 1
                    }
                }
            }
            println(out)
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_repeat_edge_zero,
    r#"
        fun main() {
            var out = 0
            repeat(0) {
                out = 99
            }
            println(out)
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_repeat_returnless,
    r#"
        fun repeatAndAdd(start: Int): Int {
            var out = start
            repeat(3) {
                out += 2
            }
            return out
        }
        fun main() {
            println(repeatAndAdd(1))
        }
    "#,
    &["7"]
);
