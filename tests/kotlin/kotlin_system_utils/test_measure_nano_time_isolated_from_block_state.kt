// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_isolated_from_block_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = "a"
            val elapsed = kotlin.system.measureNanoTime {
                value += "b"
                __check((value).toString(), "ab")
            }
            __check((value).toString(), "ab")
            __check((elapsed >= 0).toString(), "true")
        }
