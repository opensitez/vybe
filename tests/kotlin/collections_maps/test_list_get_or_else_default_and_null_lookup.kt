// vybe-test: kotlin/collections_maps/test_list_get_or_else_default_and_null_lookup
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            __check((values.getOrElse(1) { -1 }).toString(), "2")
            __check((values.getOrNull(5) ?: -1).toString(), "-1")
            __check((values.getOrElse(5) { -1 }).toString(), "-1")
            __check((values.getOrNull(0)).toString(), "1")
        }
