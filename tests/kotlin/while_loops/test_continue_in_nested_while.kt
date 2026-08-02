// vybe-test: kotlin/while_loops/test_continue_in_nested_while
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var out = 0
            while (i < 5) {
                i += 1
                var inner = 0
                while (inner < i) {
                    inner += 1
                    if (inner == 1) continue
                    out += 1
                    break
                }
            }
            println(out)
        }

