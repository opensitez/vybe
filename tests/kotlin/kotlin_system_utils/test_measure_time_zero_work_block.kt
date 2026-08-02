// vybe-test: kotlin/kotlin_system_utils/test_measure_time_zero_work_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = kotlin.system.measureTimeMillis {
                // no-op
            }
            __check((value >= 0).toString(), "true")
        }
