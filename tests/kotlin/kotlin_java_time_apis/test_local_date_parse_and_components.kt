// vybe-test: kotlin/kotlin_java_time_apis/test_local_date_parse_and_components
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalDate.parse("2024-06-15")
            __check((value.year).toString(), "2024")
            __check((value.monthValue).toString(), "6")
            __check((value.dayOfMonth).toString(), "15")
        }
