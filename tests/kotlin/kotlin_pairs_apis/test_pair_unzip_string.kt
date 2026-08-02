// vybe-test: kotlin/kotlin_pairs_apis/test_pair_unzip_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("x" to 1, "y" to 2, "z" to 3)
            val (letters, numbers) = values.unzip()
            __check((letters.joinToString("")).toString(), "xyz")
            __check((numbers.joinToString(",")).toString(), "1,2,3")
        }
