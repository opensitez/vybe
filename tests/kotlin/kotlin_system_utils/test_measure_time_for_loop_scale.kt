// vybe-test: kotlin/kotlin_system_utils/test_measure_time_for_loop_scale
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            val tiny = kotlin.system.measureTimeMillis {
                var total = 0
                for (i in 0 until 50000) {
                    total += i
                }
                println(total)
            }
            val larger = kotlin.system.measureTimeMillis {
                var total = 0L
                for (i in 0 until 100000) {
                    total += i
                }
                println(total)
            }
            println(tiny >= 0)
            println(larger >= tiny)
        }

