// vybe-test: kotlin/kotlin_java_time_apis/test_clock_system_now_monotonic_properties
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val instant = java.time.Instant.parse("2024-01-01T00:00:00Z")
            val clock = java.time.Clock.fixed(instant, java.time.ZoneId.of("UTC"))
            val a = java.time.Instant.now(clock)
            val b = java.time.Instant.now(clock)
            __check((a.toString()).toString(), "2024-01-01T00:00:00Z")
            __check((a == b).toString(), "true")
        }
