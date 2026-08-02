// vybe-test: kotlin/kotlin_map_apis/test_map_map_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val transformed = map.mapValues { it.value * 10 }
            __check((transformed["a"]).toString(), "10")
            __check((transformed["b"]).toString(), "20")
        }
