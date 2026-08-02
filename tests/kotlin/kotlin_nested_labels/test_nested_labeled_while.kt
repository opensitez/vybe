// vybe-test: kotlin/kotlin_nested_labels/test_nested_labeled_while
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_labels.rs

fun main() {
            var i = 0
            outer@ while (i < 3) {
                var j = 0
                while (j < 2) {
                    if (i == 1 && j == 0) {
                        j = j + 1
                        i = i + 1
                        continue@outer
                    }
                    println(i.toString() + ":" + j.toString())
                    j = j + 1
                }
                i = i + 1
            }
        }

