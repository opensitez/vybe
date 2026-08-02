// vybe-test: kotlin/kotlin_time/test_kotlin_time_duration_is_infinite
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Duration.INFINITE
            __check((value.isInfinite()).toString(), "true")
            __check((value.isFinite()).toString(), "false")
        }
