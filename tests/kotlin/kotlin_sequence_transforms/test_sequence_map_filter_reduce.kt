// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_map_filter_reduce
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val total = seq.map { it * 2 }
                .filter { it > 4 }
                .reduce { acc, v -> acc + v }
            __check((total).toString(), "14")
        }
