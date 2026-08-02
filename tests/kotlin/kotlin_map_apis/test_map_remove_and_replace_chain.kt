// vybe-test: kotlin/kotlin_map_apis/test_map_remove_and_replace_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("x" to 1, "y" to 2)
            map.remove("x")
            map["z"] = 3
            __check((map.size).toString(), "2")
            __check((map.containsKey("x")).toString(), "false")
            __check((map["z"]).toString(), "3")
        }
