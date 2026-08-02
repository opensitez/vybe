// vybe-test: kotlin/kotlin_java_time_apis/test_local_time_plus_hours_minutes
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = java.time.LocalTime.parse("23:45:00").plusHours(2).plusMinutes(20)
            __check((value.toString()).toString(), "02:05")
        }
