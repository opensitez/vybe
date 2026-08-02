// vybe-test: kotlin/collections/test_array_content_equals_distinguishes_reference_identity
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = arrayOf(arrayOf(1), arrayOf(2))
            val same = arrayOf(arrayOf(1), arrayOf(2))
            val deepA = left.contentDeepEquals(same)
            val sameRef = left.contentEquals(same)
            __check((deepA).toString(), "true")
            __check((sameRef).toString(), "false")
        }
