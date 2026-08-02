// vybe-test: kotlin/kotlin_java_time_apis/test_local_date_compare_to
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.time.LocalDate.parse("2024-01-01")
            val b = java.time.LocalDate.parse("2024-01-02")
            __check((a.isBefore(b)).toString(), "true")
            __check((a.isAfter(b)).toString(), "false")
        }
