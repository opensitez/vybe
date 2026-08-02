// vybe-test: kotlin/literals/test_multiline_raw_string_basic
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun main() {
            val block = """
line one
line two
"""
            println(block.trimIndent())
            println(block.lines().size)
        }

