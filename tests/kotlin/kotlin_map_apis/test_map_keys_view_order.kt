// vybe-test: kotlin/kotlin_map_apis/test_map_keys_view_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("b" to 2, "a" to 1, "c" to 3)
            __check((map.keys.joinToString(",")).toString(), "b,a,c")
            __check((map.values.joinToString(",")).toString(), "2,1,3")
        }
