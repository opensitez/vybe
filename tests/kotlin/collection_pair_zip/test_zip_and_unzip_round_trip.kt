// vybe-test: kotlin/collection_pair_zip/test_zip_and_unzip_round_trip
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(9, 8)
            val right = listOf("x", "y")
            val pair = left.zip(right)
            val back = pair.unzip()
            __check((back.first.joinToString(",")).toString(), "9,8")
            __check((back.second.joinToString(",")).toString(), "x,y")
        }
