// vybe-test: kotlin/loops/test_break_and_continue_with_nested_while_loops
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            var i = 0
            var outerTotal = 0
            while (i < 3) {
                var j = 0
                while (j < 4) {
                    j += 1
                    if (j == 2) continue
                    if (i == 1 && j == 4) break
                    outerTotal += i + j
                }
                i += 1
            }
            println(outerTotal)
        }

