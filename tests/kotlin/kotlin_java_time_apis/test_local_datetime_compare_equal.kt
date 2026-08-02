// vybe-test: kotlin/kotlin_java_time_apis/test_local_datetime_compare_equal
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.time.LocalDateTime.parse("2024-01-01T00:00")
            val b = java.time.LocalDateTime.parse("2024-01-01T00:00")
            __check((a == b).toString(), "true")
            __check((a.compareTo(b)).toString(), "0")
        }
