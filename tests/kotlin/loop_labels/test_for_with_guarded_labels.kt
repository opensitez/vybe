// vybe-test: kotlin/loop_labels/test_for_with_guarded_labels
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

fun isAllowed(x: Int): Boolean = x % 2 == 0
        fun main() {
            var out = 0
            outer@ for (i in 1..6) {
                if (!isAllowed(i)) continue@outer
                out += i
            }
            println(out)
        }

