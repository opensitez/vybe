// vybe-test: kotlin/collection_pair_zip/test_zip_with_shorter_right_iterable
// origin: languages/kotlin/tests/kotlin/test_collection_pair_zip.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("one", "two", "three", "four")
            val nums = listOf(10)
            val zipped = words.zip(nums).joinToString(",") { "${'$'}{it.first}:${'$'}{it.second}" }
            __check((zipped).toString(), "one:10")
        }
