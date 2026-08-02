// vybe-test: kotlin/loop_labels/test_label_named_loop_in_function
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun score(x: Int): Int {
            var out = 0
            outer@ for (i in 1..x) {
                if (i == 4) break@outer
                out += i
            }
            return out
        }
        fun main() {
            println(score(6))
        }

