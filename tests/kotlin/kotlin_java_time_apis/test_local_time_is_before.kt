// vybe-test: kotlin/kotlin_java_time_apis/test_local_time_is_before
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.time.LocalTime.parse("08:00:00")
            val b = java.time.LocalTime.parse("08:00:01")
            __check((a.isBefore(b)).toString(), "true")
            __check((a == b.minusSeconds(1)).toString(), "true")
        }
