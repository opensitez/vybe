// vybe-test: kotlin/kotlin_system_utils/test_measure_time_with_exception_propagation
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun main() {
            try {
                kotlin.system.measureTimeMillis {
                    throw IllegalStateException("x")
                }
                println("ok")
            } catch (e: IllegalStateException) {
                println(e.message)
            }
        }

