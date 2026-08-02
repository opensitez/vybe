// vybe-test: kotlin/kotlin_progressions/test_progression_first_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 10 downTo 2 step 4
            __check((r.first).toString(), "10")
            __check((r.last).toString(), "2")
            __check((r.step).toString(), "-4")
        }
