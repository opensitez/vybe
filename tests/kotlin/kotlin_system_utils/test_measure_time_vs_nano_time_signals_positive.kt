// vybe-test: kotlin/kotlin_system_utils/test_measure_time_vs_nano_time_signals_positive
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            val millis = kotlin.system.measureTimeMillis { for (i in 0 until 3000) {} }
            val nanos = kotlin.system.measureNanoTime { for (i in 0 until 3000) {} }
            println(millis >= 0)
            println(nanos >= 0)
            println((nanos / 1000000) >= millis)
        }

