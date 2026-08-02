// vybe-test: kotlin/type_cast_edges/test_require_cast_message
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun main() {
            val value: Any = 1
            try {
                val text = value as String
                println(text)
            } catch (e: ClassCastException) {
                println("bad_cast")
            }
        }

