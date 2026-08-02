// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_runs_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var seen = false
            val nanos = kotlin.system.measureNanoTime {
                seen = true
                val out = "x" + "y"
                __check((out).toString(), "xy")
            }
            __check((seen).toString(), "true")
            __check((nanos >= 0).toString(), "true")
        }
