// vybe-test: kotlin/try_catch_flow/test_try_nested_loop_and_error
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            var out = 0
            for (i in 1..3) {
                try {
                    if (i == 2) throw Exception("x")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
            }
            println(out)
        }

