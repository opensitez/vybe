// vybe-test: kotlin/strings/test_string_char_at_invalid_index_throws
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val value = "kotlin"
            try {
                println(value[10])
            } catch (e: java.lang.StringIndexOutOfBoundsException) {
                println("out_of_bounds")
            }
        }

