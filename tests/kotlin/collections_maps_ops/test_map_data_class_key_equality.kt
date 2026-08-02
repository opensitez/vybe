// vybe-test: kotlin/collections_maps_ops/test_map_data_class_key_equality
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

data class Key(val id: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf(Key(1) to "first", Key(2) to "second")
            __check((map[Key(1)]).toString(), "first")
            __check((map[Key(2)]).toString(), "second")
        }
