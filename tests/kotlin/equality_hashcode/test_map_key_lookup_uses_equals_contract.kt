// vybe-test: kotlin/equality_hashcode/test_map_key_lookup_uses_equals_contract
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Key(val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val cache = mapOf(Key("a") to 1)
            __check((cache[Key("a")]).toString(), "1")
            __check((cache.containsKey(Key("x"))).toString(), "false")
}
