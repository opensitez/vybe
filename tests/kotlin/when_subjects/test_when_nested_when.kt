// vybe-test: kotlin/when_subjects/test_when_nested_when
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 5
            val out = when (a) {
                in 1..10 -> when (a % 2) {
                    0 -> "even"
                    else -> "odd"
                }
                else -> "none"
            }
            __check((out).toString(), "odd")
        }
