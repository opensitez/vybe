// vybe-test: kotlin/runtime_type_queries/test_smart_cast_after_is
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

fun main() {
            val value: Any = "abc"
            if (value is String) {
                println(value.length)
            } else {
                println(0)
            }
            val boxed: Any = 123
            if (boxed is Int) {
                println(boxed + 1)
            }
        }

