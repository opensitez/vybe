// vybe-test: kotlin/collection_pair_zip/test_unzip_restores_components
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(1 to "a", 2 to "b", 3 to "c")
            val (nums, chars) = source.unzip()
            __check((nums.joinToString(",")).toString(), "1,2,3")
            __check((chars.joinToString(",")).toString(), "a,b,c")
        }
