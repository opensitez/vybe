// vybe-test: kotlin/kotlin_set_apis/test_set_hash_code_stable_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf("x", "y", "z")
            __check((set.size).toString(), "3")
            __check((set.contains("y")).toString(), "true")
        }
