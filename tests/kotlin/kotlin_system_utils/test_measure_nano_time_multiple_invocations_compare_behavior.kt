// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_multiple_invocations_compare_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            val first = kotlin.system.measureNanoTime { for (i in 1..2000) {} }
            val second = kotlin.system.measureNanoTime { for (i in 1..2000) {} }
            println(first >= 0)
            println(second >= 0)
        }

