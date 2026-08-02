// vybe-test: kotlin/collection_fold_scan/test_fold_order_non_commutative
// origin: languages/kotlin/tests/kotlin/test_collection_fold_scan.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            val left = values.fold("") { acc, value -> "${'$'}acc${'$'}{value}" }
            val right = values.foldRight("") { value, acc -> "${'$'}value${'$'}{acc}" }
            __check((left).toString(), "abc")
            __check((right).toString(), "abc")
        }
