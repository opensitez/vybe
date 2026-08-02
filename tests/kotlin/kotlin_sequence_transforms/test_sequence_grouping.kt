// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_grouping
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4)
            val grouped = seq.groupBy { it % 2 == 0 }
            val evens = grouped[true]?.joinToString(",") ?: "none"
            val odds = grouped[false]?.joinToString(",") ?: "none"
            __check((evens).toString(), "2,4")
            __check((odds).toString(), "1,3")
        }
