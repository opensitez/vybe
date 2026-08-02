// vybe-test: kotlin/kotlin_map_apis/test_map_keys_to_set_size
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            __check((map.keys.toSet().size).toString(), "2")
            __check((map.keys.toSet().toList().joinToString(",")).toString(), "a,b")
        }
