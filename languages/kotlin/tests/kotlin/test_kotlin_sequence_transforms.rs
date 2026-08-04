kotlin_run_cases! {
    test_sequence_map_filter_reduce => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val total = seq.map { it * 2 }
                .filter { it > 4 }
                .reduce { acc, v -> acc + v }
            println(total)
        }
    "##, vec![String::from("14")]),
    test_sequence_count_take => (r##"
        fun main() {
            val seq = (1..10).asSequence()
            println(seq.count())
            println(seq.take(3).toList().size)
        }
    "##, vec![String::from("10"), String::from("3")]),
    test_generate_sequence_finite => (r##"
        fun main() {
            val seq = generateSequence(1) { if (it < 4) it + 1 else null }
            println(seq.toList().joinToString(","))
        }
    "##, vec![String::from("1,2,3,4")]),
    test_sequence_sum_fold => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3)
            var sum = 0
            for (v in seq) { sum += v }
            println(sum)
            val total = sequenceOf(10, 20, 30).fold(0) { acc, value -> acc + value }
            println(total)
        }
    "##, vec![String::from("6"), String::from("60")]),
    test_sequence_any_all => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            println(seq.any { it == 3 }.toString())
            println(seq.all { it > 0 }.toString())
        }
    "##, vec![String::from("true"), String::from("true")]),
    test_sequence_join_to_string => (r##"
        fun main() {
            val seq = sequenceOf("a", "b", "c")
            println(seq.joinToString("-"))
        }
    "##, vec![String::from("a-b-c")]),
    test_sequence_zip => (r##"
        fun main() {
            val a = sequenceOf(1, 2, 3)
            val b = sequenceOf("x", "y", "z")
            val zipped = a.zip(b)
            println(zipped.toList().joinToString(",") { (n, s) -> "$n$s" })
        }
    "##, vec![String::from("1x,2y,3z")]),
    test_sequence_flat_map => (r##"
        fun main() {
            val seq = sequenceOf(1, 2).flatMap { v -> sequenceOf(v, v * 10) }
            println(seq.toList().joinToString(","))
        }
    "##, vec![String::from("1,10,2,20")]),
    test_sequence_take_while => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 0, 4)
            println(seq.takeWhile { it > 0 }.toList().joinToString(","))
            println(seq.dropWhile { it < 3 }.toList().joinToString(","))
        }
    "##, vec![String::from("1,2,3"), String::from("3,0,4")]),
    test_sequence_transform_with_index => (r##"
        fun main() {
            val seq = sequenceOf("a", "b", "c")
            val withIndex = seq.withIndex().map { "${it.index}:${it.value}" }
            println(withIndex.toList().joinToString(","))
        }
    "##, vec![String::from("0:a,1:b,2:c")]),
    test_sequence_element_at => (r##"
        fun main() {
            val seq = sequenceOf(5, 6, 7)
            println(seq.elementAt(1).toString())
            println(seq.elementAtOrElse(10) { -1 }.toString())
        }
    "##, vec![String::from("6"), String::from("-1")]),
    test_sequence_partition => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4, 5)
            val (evens, odds) = seq.partition { it % 2 == 0 }
            println(evens.toList().joinToString(","))
            println(odds.toList().joinToString(","))
        }
    "##, vec![String::from("2,4"), String::from("1,3,5")]),
    test_sequence_for_each => (r##"
        fun main() {
            var total = 0
            sequenceOf(1, 2, 3).forEach { total += it }
            println(total)
        }
    "##, vec![String::from("6")]),
    test_sequence_to_mutable => (r##"
        fun main() {
            val list = sequenceOf(1, 2, 3).toMutableList()
            list.add(4)
            println(list.joinToString(","))
        }
    "##, vec![String::from("1,2,3,4")]),
    test_sequence_grouping => (r##"
        fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val grouped = seq.groupBy { it % 2 == 0 }
            val evens = grouped[true]?.joinToString(",") ?: "none"
            val odds = grouped[false]?.joinToString(",") ?: "none"
            println(evens)
            println(odds)
        }
    "##, vec![String::from("2,4"), String::from("1,3")]),
    test_sequence_replayability => (r##"
        fun main() {
            val seq = generateSequence(1) { if (it < 3) it + 1 else null }
            println(seq.toList().size)
            // recreate because sequences are consumed
            val seq2 = generateSequence(1) { if (it < 3) it + 1 else null }
            println(seq2.toList().sum().toString())
        }
    "##, vec![String::from("3"), String::from("6")]),
}
