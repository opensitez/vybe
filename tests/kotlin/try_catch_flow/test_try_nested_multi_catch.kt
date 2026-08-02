// vybe-test: kotlin/try_catch_flow/test_try_nested_multi_catch
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: RuntimeException) {
                println("runtime")
            } catch (e: Exception) {
                println("general")
            }
        }

