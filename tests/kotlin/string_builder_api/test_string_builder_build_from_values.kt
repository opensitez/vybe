// vybe-test: kotlin/string_builder_api/test_string_builder_build_from_values
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun main() {
            val out = StringBuilder()
            val a = listOf("x", "y", "z")
            for (item in a) {
                out.append(item)
            }
            println(out.toString())
        }

