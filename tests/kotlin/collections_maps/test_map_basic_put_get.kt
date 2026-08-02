// vybe-test: kotlin/collections_maps/test_map_basic_put_get
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("alice" to 3, "bob" to 7)
            __check((scores["alice"]).toString(), "3")
            __check((scores["bob"]).toString(), "7")
            __check((scores.size).toString(), "2")
        }
