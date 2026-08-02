// vybe-test: kotlin/collection_pair_zip/test_product_via_zip
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names = listOf("a", "b", "c")
            val out = names.zip(generateSequence(0) { it + 1 }) { name, i -> "${'$'}name${'$'}i" }
            __check((out.joinToString(",")).toString(), "a0,b1,c2")
        }
