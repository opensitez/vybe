// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_nested_blocks
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val outer = kotlin.system.measureNanoTime {
                kotlin.system.measureNanoTime {
                    __check(("inner").toString(), "inner")
                }
            }
            __check((outer >= 0).toString(), "true")
        }
