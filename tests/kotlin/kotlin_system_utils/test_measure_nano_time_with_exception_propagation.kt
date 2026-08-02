// vybe-test: kotlin/kotlin_system_utils/test_measure_nano_time_with_exception_propagation
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            try {
                kotlin.system.measureNanoTime {
                    throw IllegalStateException("n")
                }
                println("ok")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }

