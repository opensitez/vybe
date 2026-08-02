// vybe-test: kotlin/while_loops/test_while_post_condition
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 5
            var out = 0
            while (i >= 1) {
                out += i
                i -= 2
            }
            println(out)
        }

