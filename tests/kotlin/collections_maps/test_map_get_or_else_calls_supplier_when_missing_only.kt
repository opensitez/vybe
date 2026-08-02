// vybe-test: kotlin/collections_maps/test_map_get_or_else_calls_supplier_when_missing_only
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("a" to 1)
            var asked = 0
            val miss = scores.getOrElse("b") {
                asked += 1
                99
            }
            val hit = scores.getOrElse("a") {
                asked += 1
                88
            }
            __check((miss).toString(), "99")
            __check((hit).toString(), "1")
            __check((asked).toString(), "1")
        }
