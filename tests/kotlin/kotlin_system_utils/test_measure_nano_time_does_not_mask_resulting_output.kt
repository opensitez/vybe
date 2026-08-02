// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_does_not_mask_resulting_output
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                val left = kotlin.system.measureTimeMillis { __check((1).toString(), "1") }
                val right = kotlin.system.measureTimeMillis { __check((2).toString(), "2") }
                __check((left + right).toString(), "3")
            }
            __check((elapsed >= 0).toString(), "true")
        }
