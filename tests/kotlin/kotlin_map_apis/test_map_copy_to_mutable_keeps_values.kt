// vybe-test: kotlin/kotlin_map_apis/test_map_copy_to_mutable_keeps_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = linkedMapOf("a" to 1, "b" to 2)
            val copy = source.toMutableMap()
            copy.put("c", 3)
            __check((source.size).toString(), "2")
            __check((copy.size).toString(), "3")
            __check((copy["c"]).toString(), "3")
        }
