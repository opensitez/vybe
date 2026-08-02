// vybe-test: kotlin/when_subjects/test_when_guard_condition
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Pair(2, 4)
            val out = when {
                p.first == p.second -> "eq"
                p.first + p.second == 6 -> "sum-six"
                else -> "other"
            }
            __check((out).toString(), "sum-six")
        }
