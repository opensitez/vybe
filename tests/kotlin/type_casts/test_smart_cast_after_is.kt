// vybe-test: kotlin/type_casts/test_smart_cast_after_is
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val input: Any = 42
            if (input is Int) {
                val n: Int = input
                println(n + 1)
            } else {
                println(0)
            }
        }

