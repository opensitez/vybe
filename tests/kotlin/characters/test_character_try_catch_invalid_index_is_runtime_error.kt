// vybe-test: kotlin/characters/test_character_try_catch_invalid_index_is_runtime_error
// origin: languages/kotlin/tests/kotlin/test_characters.rs

fun main() {
            val value = "ok"
            try {
                println(value[9])
            } catch (e: Exception) {
                println("out-of-range")
            }
        }

