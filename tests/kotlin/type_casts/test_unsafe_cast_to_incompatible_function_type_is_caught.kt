// vybe-test: kotlin/type_casts/test_unsafe_cast_to_incompatible_function_type_is_caught
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val handler: Any = { value: String -> value + "!" }
            try {
                val bad = handler as (Int) -> String
                println("bad:" + bad(3))
            } catch (e: Exception) {
                println("caught")
            }
        }

