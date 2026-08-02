// vybe-test: kotlin/kotlin_pairs_apis/test_pair_list_destructure_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun main() {
            val pairs = listOf("a" to 1, "b" to 2)
            var out = ""
            for ((k, v) in pairs) {
                out += k + v.toString()
            }
            println(out)
        }

