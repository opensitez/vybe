// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_when_on_map_entries
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mapOf("a" to 3, "b" to 0)
            val out = values.map { (k, v) ->
                when (v) {
                    0 -> k + "zero"
                    else -> k + "nz"
                }
            }
            __check((out.joinToString(",")).toString(), "azero,bnz")
        }
