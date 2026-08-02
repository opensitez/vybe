// vybe-test: kotlin/strings/test_raw_string_multiline_and_indentation_removal
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val value = """
                line one
                line two
            """.trimIndent()
            println(value)
            println(value.lines().size)
        }

