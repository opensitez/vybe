// vybe-test: kotlin/while_loops/test_do_while_executes_once_when_false
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            val out = do {
                i += 1
                i
            } while (i > 10)
            println(out)
        }

