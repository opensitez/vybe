// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_large_block_is_still_non_negative
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                var value = 0L
                repeat(12000) {
                    value += it.toLong()
                }
                __check((value).toString(), "71994000")
            }
            __check((elapsed >= 0).toString(), "true")
        }
