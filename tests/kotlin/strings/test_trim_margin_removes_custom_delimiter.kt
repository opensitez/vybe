// vybe-test: kotlin/strings/test_trim_margin_removes_custom_delimiter
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun main() {
            val value = """
                |k
                |otlin
                """.trimMargin("|")
            println(value)
            println(value.lines().size)
        }

