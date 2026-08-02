// vybe-test: kotlin/kotlin_java_time_apis/test_epoch_millis_roundtrip_with_instant
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val instant = java.time.Instant.ofEpochMilli(1_700_000_000_000)
            __check((instant.epochSecond).toString(), "1700000000")
            __check((java.time.Instant.ofEpochSecond(instant.epochSecond).toEpochMilli() >= 1_700_000_000_000).toString(), "true")
        }
