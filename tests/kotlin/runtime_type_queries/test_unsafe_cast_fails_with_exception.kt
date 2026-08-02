// vybe-test: kotlin/runtime_type_queries/test_unsafe_cast_fails_with_exception
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun main() {
            val a: Any = 7
            try {
                val s = a as String
                println(s)
            } catch (e: Exception) {
                println("err")
            }
        }

