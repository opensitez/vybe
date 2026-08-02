// vybe-test: kotlin/kotlin_java_time_apis/test_instant_duration_between_instant_points
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.time.Instant.parse("2024-01-01T00:00:00Z")
            val b = java.time.Instant.parse("2024-01-01T00:00:30Z")
            val d = java.time.Duration.between(a, b)
            __check((d.seconds).toString(), "30")
            __check((d.toMillis()).toString(), "30000")
        }
