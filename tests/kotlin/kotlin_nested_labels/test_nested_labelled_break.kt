// vybe-test: kotlin/kotlin_nested_labels/test_nested_labelled_break
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_labels.rs

fun main() {
            outer@ for (i in 1..3) {
                for (j in 1..3) {
                    if (i == 2 && j == 2) {
                        break@outer
                    }
                    println(i.toString() + "-" + j.toString())
                }
            }
        }

