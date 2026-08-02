// vybe-test: kotlin/tuples/test_tuple_pair_as_map_key_lookup_with_identity_by_value
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val key = Pair("id", 3)
            val map = mapOf(key to "found")
            __check((map[Pair("id", 3)]).toString(), "found")
            __check((map[Pair("id", 4)] ?: "missing").toString(), "missing")
        }
