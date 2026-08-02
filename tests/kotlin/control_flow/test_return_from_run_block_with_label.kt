// vybe-test: kotlin/control_flow/test_return_from_run_block_with_label
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val found = run {
                for (i in 1..4) {
                    if (i == 3) {
                        return@run i
                    }
                }
                0
            }
            println(found)
        }

