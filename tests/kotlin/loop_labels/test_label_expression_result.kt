// vybe-test: kotlin/loop_labels/test_label_expression_result
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun main() {
            val value = run {
                outer@ for (i in 1..3) {
                    if (i == 2) continue@outer
                }
                42
            }
            println(value)
        }

