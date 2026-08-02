// vybe-test: kotlin/kotlin_map_apis/test_map_retain_order_after_update
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            map["b"] = 9
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map["b"]).toString(), "9")
        }
