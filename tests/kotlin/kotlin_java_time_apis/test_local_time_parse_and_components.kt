// vybe-test: kotlin/kotlin_java_time_apis/test_local_time_parse_and_components
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalTime.parse("09:30:45")
            __check((value.hour).toString(), "9")
            __check((value.minute).toString(), "30")
            __check((value.second).toString(), "45")
        }
