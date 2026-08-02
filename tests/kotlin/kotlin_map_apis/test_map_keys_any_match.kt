// vybe-test: kotlin/kotlin_map_apis/test_map_keys_any_match
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("alpha" to 1, "beta" to 2, "gamma" to 3)
            val hasLong = map.keys.any { it.length > 4 }
            val hasShort = map.keys.none { it.length > 6 }
            __check((hasLong).toString(), "true")
            __check((hasShort).toString(), "true")
        }
