// vybe-test: kotlin/kotlin_java_time_apis/test_duration_parse_and_multiply
// origin: languages/kotlin/tests/kotlin/test_kotlin_java_time_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val d = java.time.Duration.parse("PT2H30M")
            __check((d.toHours()).toString(), "2")
            __check((d.minusHours(1).toMinutes()).toString(), "90")
        }
