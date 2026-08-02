// vybe-test: kotlin/java_util_collections/test_java_collections_n_copies_materializes_repeated_value
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Collections.nCopies(4, "zap")
            __check((values.size).toString(), "4")
            __check((values[0]).toString(), "zap")
            __check((values[3]).toString(), "zap")
        }
