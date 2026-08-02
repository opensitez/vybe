// vybe-test: kotlin/kotlin_time/test_kotlin_time_zero_plus_zero_stable
// origin: languages/kotlin/tests/kotlin/test_kotlin_time.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Duration.ZERO + Duration.ZERO
            __check((value == Duration.ZERO).toString(), "true")
            __check((value.inWholeMilliseconds).toString(), "0")
        }
