// vybe-test: kotlin/collections_maps_ops/test_map_get_or_else_with_supplier
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("one" to 1)
            val existing = map.getOrElse("one") { 99 }
            val missing = map.getOrElse("two") { 99 }
            __check((existing).toString(), "1")
            __check((missing).toString(), "99")
        }
