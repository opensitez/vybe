// vybe-test: kotlin/boolean_logic/test_boolean_operator_normalization_in_loops
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun main() {
            var i = 0
            var ok = true
            while (ok && i < 3) {
                i++
                if (i == 2) {
                    ok = false
                }
            }
            println(i)
            println(ok)
        }

