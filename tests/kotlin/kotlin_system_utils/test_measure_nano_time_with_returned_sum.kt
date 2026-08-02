// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_with_returned_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = kotlin.system.measureNanoTime {
                2 + 3
            }
            __check((result >= 0).toString(), "true")
            __check((kotlin.system.measureTimeMillis {
                7 * 6
            } >= 0).toString(), "true")
        }
