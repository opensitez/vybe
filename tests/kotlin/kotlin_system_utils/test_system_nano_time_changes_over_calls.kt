// vybe-test: kotlin/kotlin_system_utils/test_system_nano_time_changes_over_calls
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = System.nanoTime()
            val second = System.nanoTime()
            __check((second >= first).toString(), "true")
        }
