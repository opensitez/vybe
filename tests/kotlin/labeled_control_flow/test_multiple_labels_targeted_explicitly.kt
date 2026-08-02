// vybe-test: kotlin/labeled_control_flow/test_multiple_labels_targeted_explicitly
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var x = 0
            outer@ for (i in 0..1) {
                inner@ for (j in 0..2) {
                    if (i == 1 && j == 0) {
                        continue@outer
                    }
                    if (j == 2) {
                        break@inner
                    }
                    x += 1
                }
            }
            println(x)
        }

