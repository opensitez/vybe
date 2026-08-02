// vybe-test: kotlin/kotlin_java_time_apis/test_duration_plus_and_minus
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = java.time.Duration.ofHours(1).plusMinutes(90).minusMinutes(15)
            __check((d.toMinutes()).toString(), "135")
        }
