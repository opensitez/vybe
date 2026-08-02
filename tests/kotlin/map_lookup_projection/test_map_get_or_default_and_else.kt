// vybe-test: kotlin/map_lookup_projection/test_map_get_or_default_and_else
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("one" to 1, "two" to 2)
            __check((source.getOrDefault("two", -1)).toString(), "2")
            __check((source.getOrDefault("x", -1)).toString(), "-1")
            __check((source.getOrElse("two") { 0 }).toString(), "2")
            __check((source.getOrElse("x") { 7 }).toString(), "7")
        }
