// vybe-test: kotlin/kotlin_system_utils/test_measure_time_with_captured_return_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = kotlin.system.measureTimeMillis {
                9 + 1
            }
            val value = result / kotlin.system.measureTimeMillis {
                1
            }
            __check((value >= 0).toString(), "true")
        }
