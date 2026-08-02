// vybe-test: kotlin/when_subjects/test_when_subject_block_with_side_effect
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

var seen = 0
        fun classify(x: Int): String {
            return when (x) {
                1 -> { seen = 1
"one" }
                2 -> { seen = 2
"two" }
                else -> { seen = 3
"other" }
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(2)).toString(), "two")
            __check((seen).toString(), "2")
        }
