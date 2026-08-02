// vybe-test: kotlin/kotlin_pairs_apis/test_pair_to_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = "x" to 1
            __check((p.toString()).toString(), "(x, 1)")
        }
