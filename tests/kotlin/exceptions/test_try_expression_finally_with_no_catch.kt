// vybe-test: kotlin/exceptions/test_try_expression_finally_with_no_catch
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun status(flag: Boolean): String {
            return try {
                if (flag) "ok" else throw Exception("bad")
            } finally {
                println("cleanup")
            }
        }

        fun main() {
            try {
                println(status(false))
            } catch (e: Exception) {
                println(e.message)
            }
        }

