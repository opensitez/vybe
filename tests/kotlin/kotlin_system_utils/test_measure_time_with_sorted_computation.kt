// vybe-test: kotlin/kotlin_system_utils/test_measure_time_with_sorted_computation
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val numbers = (1..20).toList().shuffled().sorted()
            val elapsed = kotlin.system.measureTimeMillis {
                __check((numbers.joinToString(",")).toString(), "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20")
            }
            __check((numbers.size).toString(), "20")
            __check((elapsed >= 0).toString(), "true")
        }
