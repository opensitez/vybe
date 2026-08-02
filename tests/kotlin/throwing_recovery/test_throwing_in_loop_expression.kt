// vybe-test: kotlin/throwing_recovery/test_throwing_in_loop_expression
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            var out = 0
            var i = 0
            do {
                try {
                    if (i == 2) throw Exception("x")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
                i += 1
            } while (i < 4)
            println(out)
        }

