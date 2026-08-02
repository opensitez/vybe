// vybe-test: kotlin/type_casts/test_unsafe_cast_to_wrong_type_is_caught
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            try {
                val value: Any = true
                val number = value as Int
                println(number)
            } catch (e: Exception) {
                println("caught")
            }
        }

