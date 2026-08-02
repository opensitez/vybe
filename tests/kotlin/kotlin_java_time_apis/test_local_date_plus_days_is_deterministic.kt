// vybe-test: kotlin/kotlin_java_time_apis/test_local_date_plus_days_is_deterministic
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-12-30").plusDays(5)
            __check((value.toString()).toString(), "2025-01-04")
        }
