// vybe-test: kotlin/kotlin_java_time_apis/test_local_datetime_split_and_combine
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDateTime.parse("2024-07-01T10:11:12")
            __check((value.date.toString()).toString(), "2024-07-01")
            __check((value.toLocalTime().toString()).toString(), "10:11:12")
        }
