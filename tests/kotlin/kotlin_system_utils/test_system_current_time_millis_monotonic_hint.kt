// vybe-test: kotlin/kotlin_system_utils/test_system_current_time_millis_monotonic_hint
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = System.currentTimeMillis()
            val b = System.currentTimeMillis()
            __check((b >= a).toString(), "true")
        }
