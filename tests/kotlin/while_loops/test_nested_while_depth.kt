// vybe-test: kotlin/while_loops/test_nested_while_depth
// origin: languages/kotlin/tests/kotlin/test_while_loops.rs

fun main() {
            var i = 0
            var total = 0
            while (i < 3) {
                var j = 0
                while (j < 2) {
                    total += i + j
                    j += 1
                }
                i += 1
            }
            println(total)
        }

