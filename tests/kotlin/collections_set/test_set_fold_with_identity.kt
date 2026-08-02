// vybe-test: kotlin/collections_set/test_set_fold_with_identity
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            __check((values.fold(0) { acc, item -> acc + item }).toString(), "6")
            __check((values.reduce { acc, item -> acc * item }).toString(), "6")
        }
