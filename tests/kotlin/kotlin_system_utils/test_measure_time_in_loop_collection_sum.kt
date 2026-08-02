// vybe-test: kotlin/kotlin_system_utils/test_measure_time_in_loop_collection_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            val times = mutableListOf<Long>()
            repeat(3) {
                val elapsed = kotlin.system.measureTimeMillis {
                    var sum = 0
                    for (i in 1..5000) sum += i
                    if (sum == 0) println("x")
                }
                times.add(elapsed)
            }
            println(times.size)
            println(times.all { it >= 0 })
        }

