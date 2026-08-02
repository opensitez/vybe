// vybe-test: kotlin/try_catch_flow/test_nested_try_no_throw
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                val ok = try {
                    2 + 2
                } catch (e: Exception) {
                    -1
                }
                println(ok)
            } catch (e: Exception) {
                println("outer")
            }
        }

