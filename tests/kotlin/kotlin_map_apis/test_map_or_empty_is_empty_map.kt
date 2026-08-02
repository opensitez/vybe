// vybe-test: kotlin/kotlin_map_apis/test_map_or_empty_is_empty_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map: Map<String, Int> = mapOf()
            val empty = map.orEmpty()
            __check((empty.isEmpty()).toString(), "true")
            __check((empty.size).toString(), "0")
        }
