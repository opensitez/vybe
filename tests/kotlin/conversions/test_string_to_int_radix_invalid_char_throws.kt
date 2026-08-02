// vybe-test: kotlin/conversions/test_string_to_int_radix_invalid_char_throws
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun main() {
            try {
                println("2".toInt(2))
            } catch (e: NumberFormatException) {
                println("bad")
            }
        }

