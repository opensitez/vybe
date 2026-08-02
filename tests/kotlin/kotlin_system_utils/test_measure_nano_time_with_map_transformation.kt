// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_with_map_transformation
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf(1, 2, 3)
            val elapsed = kotlin.system.measureNanoTime {
                val out = items.map { it * 2 }
                __check((out.joinToString(",")).toString(), "2,4,6")
            }
            __check((elapsed >= 0).toString(), "true")
        }
