// vybe-test: kotlin/literals/test_raw_string_with_margin
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun main() {
            val block = """
                >a
                >b
            """.trimMargin(">")
            println(block)
            println(block.lines().size)
        }

