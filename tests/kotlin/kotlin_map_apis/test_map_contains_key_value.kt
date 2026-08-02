// vybe-test: kotlin/kotlin_map_apis/test_map_contains_key_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            __check((map.containsKey("a")).toString(), "true")
            __check((map.containsKey("z")).toString(), "false")
            __check((map.containsValue(2)).toString(), "true")
        }
