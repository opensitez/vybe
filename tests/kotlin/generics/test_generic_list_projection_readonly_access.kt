// vybe-test: kotlin/generics/test_generic_list_projection_readonly_access
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> joinValues(values: List<out T>): String {
            var out = ""
            for (value in values) {
                out += value.toString()
            }
            return out
        }

        fun main() {
            val ints = listOf(1, 2, 3)
            val texts: List<String> = listOf("a", "b")
            println(joinValues(ints))
            println(joinValues(texts))
        }

