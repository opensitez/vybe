// vybe-test: kotlin/java_util_collections/test_java_collections_n_copies_contains_every_copy
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Collections.nCopies(3, 7)
            __check((values.contains(7)).toString(), "true")
            __check((values.contains(2)).toString(), "false")
            __check((values.indexOf(7)).toString(), "0")
        }
