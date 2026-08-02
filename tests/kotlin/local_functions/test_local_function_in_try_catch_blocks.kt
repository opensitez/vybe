// vybe-test: kotlin/local_functions/test_local_function_in_try_catch_blocks
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun main() {
            try {
                fun parse(v: String): Int {
                    if (v.length == 0) throw RuntimeException("empty")
                    return v.toInt()
                }
                println(parse("12"))
            } catch (error: RuntimeException) {
                println("bad")
            }
        }

