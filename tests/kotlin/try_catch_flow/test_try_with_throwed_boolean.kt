// vybe-test: kotlin/try_catch_flow/test_try_with_throwed_boolean
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                if (true) throw Exception("x")
                println("never")
            } catch (e: Exception) {
                println(e.message)
            } finally {
                println("done")
            }
        }

