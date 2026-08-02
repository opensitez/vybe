// vybe-test: kotlin/kotlin_system_utils/test_measure_time_millis_runs_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            var seen = false
            val millis = kotlin.system.measureTimeMillis {
                seen = true
                var sum = 0
                for (i in 1..1000) sum += i
                println(sum)
            }
            println(seen)
            println(millis >= 0)
        }

