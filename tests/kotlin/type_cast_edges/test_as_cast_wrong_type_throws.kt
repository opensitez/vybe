// vybe-test: kotlin/type_cast_edges/test_as_cast_wrong_type_throws
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun main() {
            val value: Any? = 10
            try {
                value as String
                println("ok")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
        }

