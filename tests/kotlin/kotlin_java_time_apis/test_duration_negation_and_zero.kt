// vybe-test: kotlin/kotlin_java_time_apis/test_duration_negation_and_zero
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = java.time.Duration.ofMinutes(10).negated()
            __check((d.toMinutes()).toString(), "-10")
            __check(((d + java.time.Duration.ofMinutes(10)).isZero()).toString(), "true")
        }
