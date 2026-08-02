// vybe-test: kotlin/kotlin_java_time_apis/test_zoned_date_time_parse_with_zone
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.ZonedDateTime.parse("2024-01-01T10:15:30+01:00[Europe/Paris]")
            __check((value.zone.id).toString(), "Europe/Paris")
            __check((value.offset).toString(), "+01:00")
            __check((value.toLocalDateTime().toString()).toString(), "2024-01-01T10:15:30")
        }
