// vybe-test: kotlin/try_catch_flow/test_try_multiple_catch_branches
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                throw Exception("x")
            } catch (e: IllegalArgumentException) {
                println("illegal")
            } catch (e: Exception) {
                println("general")
            }
        }

