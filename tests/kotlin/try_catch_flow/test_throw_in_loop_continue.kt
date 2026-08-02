// vybe-test: kotlin/try_catch_flow/test_throw_in_loop_continue
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            var out = 0
            for (i in 0..3) {
                try {
                    if (i == 2) throw Exception("stop")
                    out += i
                } catch (e: Exception) {
                    out += 10
                }
            }
            println(out)
        }

