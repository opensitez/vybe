// vybe-test: kotlin/kotlin_map_apis/test_map_remove_missing
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1)
            val removed = map.remove("z")
            __check((removed == null).toString(), "true")
            __check((map.size).toString(), "1")
        }
