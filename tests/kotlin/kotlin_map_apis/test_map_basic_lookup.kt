// vybe-test: kotlin/kotlin_map_apis/test_map_basic_lookup
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            __check((map["a"]).toString(), "1")
            __check((map["c"]).toString(), "3")
            __check((map.size).toString(), "3")
        }
