// vybe-test: kotlin/control_flow/test_repeat_negative_count_throws
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            try {
                repeat(-1) {
                    println("bad")
                }
            } catch (e: Exception) {
                println("caught")
            }
        }

