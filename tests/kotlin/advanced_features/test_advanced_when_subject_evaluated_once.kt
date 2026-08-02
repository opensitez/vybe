// vybe-test: kotlin/advanced_features/test_advanced_when_subject_evaluated_once
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

var calls = 0

        fun tapped(): Int {
            calls += 1
            return calls
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val status = when (tapped()) {
                1 -> "first"
                2 -> "second"
                else -> "other"
            }
            __check((status).toString(), "first")
            __check((calls).toString(), "1")
        }
