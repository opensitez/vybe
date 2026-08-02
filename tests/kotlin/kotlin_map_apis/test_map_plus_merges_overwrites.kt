// vybe-test: kotlin/kotlin_map_apis/test_map_plus_merges_overwrites
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = linkedMapOf("a" to 1, "b" to 2)
            val b = linkedMapOf("b" to 9, "c" to 3)
            val merged = a + b
            __check((merged["a"]).toString(), "1")
            __check((merged["b"]).toString(), "9")
            __check((merged["c"]).toString(), "3")
        }
