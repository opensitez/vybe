// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_collection_derived_metric
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val elapsed = kotlin.system.measureNanoTime {
                val text = listOf("a", "b", "c").joinToString("")
                __check((text).toString(), "abc")
            }
            __check((elapsed >= 0).toString(), "true")
        }
