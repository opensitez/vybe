// vybe-test: kotlin/kotlin_pairs_apis/test_triple_string_form
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Triple("k", "t", "v")
            __check((t.toString()).toString(), "(k, t, v)")
        }
