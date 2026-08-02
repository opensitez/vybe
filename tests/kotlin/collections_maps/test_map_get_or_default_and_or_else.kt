// vybe-test: kotlin/collections_maps/test_map_get_or_default_and_or_else
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("a" to 10, "b" to 20)
            __check((scores.getOrDefault("a", 0)).toString(), "10")
            __check((scores.getOrDefault("c", 30)).toString(), "30")
            __check((scores.getOrElse("b") { 0 }).toString(), "20")
            __check((scores.getOrElse("c") { 99 }).toString(), "99")
        }
