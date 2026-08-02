// vybe-test: kotlin/kotlin_nested_labels/test_nested_labelled_continue
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_labels.rs

fun main() {
            outer@ for (i in 1..2) {
                for (j in 1..3) {
                    if (j == 2) {
                        continue@outer
                    }
                    println(i.toString() + "-" + j.toString())
                }
            }
        }

