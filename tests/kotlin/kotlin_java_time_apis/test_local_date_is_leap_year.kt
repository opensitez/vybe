// vybe-test: kotlin/kotlin_java_time_apis/test_local_date_is_leap_year
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((java.time.LocalDate.parse("2024-02-01").isLeapYear()).toString(), "true")
            __check((java.time.LocalDate.parse("2023-02-01").isLeapYear()).toString(), "false")
        }
