// vybe-test: kotlin/kotlin_java_time_apis/test_offset_date_time_parse
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.OffsetDateTime.parse("2024-06-01T12:00:00+02:00")
            __check((value.offset.id).toString(), "+02:00")
            __check((value.toLocalDateTime().toString()).toString(), "2024-06-01T12:00")
            __check((value.toInstant().toEpochMilli()).toString(), "1717236000000")
        }
