// vybe-test: kotlin/kotlin_system_utils/test_measure_time_large_block_is_still_non_negative
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val elapsed = kotlin.system.measureTimeMillis {
                var value = 0
                repeat(10000) {
                    value += it
                }
                __check((value).toString(), "49995000")
            }
            __check((elapsed >= 0).toString(), "true")
        }
