// vybe-test: kotlin/kotlin_java_time_apis/test_zoned_date_time_with_fixed_offset
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zone = java.time.ZoneId.of("UTC")
            val value = java.time.ZonedDateTime.of(
                java.time.LocalDateTime.of(2024, 1, 1, 12, 0),
                zone
            )
            __check((value.toOffsetDateTime().offset.id).toString(), "Z")
            __check((value.toInstant().toEpochMilli()).toString(), "1704110400000")
        }
