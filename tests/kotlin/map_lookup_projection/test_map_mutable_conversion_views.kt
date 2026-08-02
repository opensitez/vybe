// vybe-test: kotlin/map_lookup_projection/test_map_mutable_conversion_views
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val mutable = source.toMutableMap()
            mutable["c"] = 3
            mutable.remove("a")
            __check((mutable.size).toString(), "2")
            __check((mutable.keys.joinToString(",")).toString(), "b,c")
            val restored = mutable.toMap()
            __check((restored.containsKey("a")).toString(), "false")
            __check((restored["c"]).toString(), "3")
        }
