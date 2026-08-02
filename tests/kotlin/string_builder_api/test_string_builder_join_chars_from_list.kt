// vybe-test: kotlin/string_builder_api/test_string_builder_join_chars_from_list
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun main() {
            val out = StringBuilder()
            val chars = listOf('a', 'b', 'c')
            for (c in chars) out.append(c)
            println(out.toString())
        }

