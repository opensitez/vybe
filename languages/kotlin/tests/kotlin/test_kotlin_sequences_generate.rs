kotlin_run_test!(
    test_generate_sequence_take_three,
    r#"
        fun main() {
            val seq = generateSequence(1) { it + 1 }.take(3).toList()
            println(seq.joinToString(","))
        }
    "#,
    &["1,2,3"]
);

kotlin_run_test!(
    test_generate_sequence_with_null_stop,
    r#"
        var i = 0
        fun step(value: Int): Int? {
            return if (value < 4) value + 1 else null
        }

        fun main() {
            val seq = generateSequence(0) { step(it) }
            println(seq.toList().joinToString(","))
        }
    "#,
    &["1,2,3,4"]
);

kotlin_run_test!(
    test_sequence_of_literals_then_aggregate,
    r#"
        fun main() {
            val s = sequenceOf("a", "b", "c")
            println(s.joinToString("|"))
        }
    "#,
    &["a|b|c"]
);

kotlin_run_test!(
    test_as_sequence_projection_chain,
    r#"
        fun main() {
            val out = (1..6).asSequence().filter { it % 2 == 0 }.map { it * it }.take(2).toList()
            println(out.joinToString(","))
        }
    "#,
    &["4,16"]
);

kotlin_run_test!(
    test_sequence_map_and_reduce,
    r#"
        fun main() {
            val out = sequenceOf(1, 2, 3)
                .map { it + 1 }
                .reduce { acc, value -> acc + value }
            println(out)
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_sequence_with_take_while,
    r#"
        fun main() {
            val out = generateSequence(1) { it + 2 }
                .takeWhile { it < 8 }
                .toList()
            println(out.joinToString(","))
        }
    "#,
    &["1,3,5,7"]
);

kotlin_run_test!(
    test_sequence_zipwithnext_side_effect,
    r#"
        fun main() {
            val values = sequenceOf(1, 2, 3, 4).zipWithNext()
            println(values.toList().joinToString("|") { "${it.first}-${it.second}" })
        }
    "#,
    &["1-2|2-3|3-4"]
);

kotlin_run_test!(
    test_sequence_windowed_projection,
    r#"
        fun main() {
            val values = generateSequence(0) { if (it == 3) null else it + 1 }
                .windowed(2)
                .toList()
            println(values.size)
            println(values.joinToString("|") { it.joinToString("") })
        }
    "#,
    &["3", "01|12|23"]
);

kotlin_run_test!(
    test_sequence_chunked_with_transform,
    r#"
        fun main() {
            val sum = (1..6).asSequence().chunked(2) { it.sum() }.toList()
            println(sum.joinToString(","))
        }
    "#,
    &["3,7,11"]
);

kotlin_run_test!(
    test_iterator_style_consumption_chain,
    r#"
        fun main() {
            val source = sequenceOf("a", "bb", "c").iterator()
            var out = ""
            while (source.hasNext()) {
                out += source.next()
            }
            println(out)
        }
    "#,
    &["abbc"]
);

kotlin_run_test!(
    test_infinite_sequence_is_not_eager,
    r#"
        var calls = 0
        fun main() {
            val values = generateSequence(1) { calls++; it + 1 }
            val taken = values.take(4).toList()
            println(taken.joinToString(","))
            println(calls >= 1)
        }
    "#,
    &["1,2,3,4", "true"]
);

kotlin_run_test!(
    test_sequence_from_function_with_initial_state,
    r#"
        fun main() {
            val values = sequence {
                var i = 0
                while (i < 3) {
                    yield(i)
                    i += 1
                }
            }
            println(values.map { it + 1 }.joinToString(","))
        }
    "#,
    &["1,2,3"]
);
