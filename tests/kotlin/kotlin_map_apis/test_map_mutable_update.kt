// vybe-test: kotlin/kotlin_map_apis/test_map_mutable_update
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            map["a"] = 8
            map["c"] = 4
            __check((map["a"]).toString(), "8")
            __check((map["c"]).toString(), "4")
            __check((map.size).toString(), "3")
        }
