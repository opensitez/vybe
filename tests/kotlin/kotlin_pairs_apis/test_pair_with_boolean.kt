// vybe-test: kotlin/kotlin_pairs_apis/test_pair_with_boolean
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun main() {
            val values = listOf(true to "on", false to "off")
            var out = ""
            for ((state, name) in values) {
                out += if (state) name.uppercase() else name
            }
            println(out)
        }

