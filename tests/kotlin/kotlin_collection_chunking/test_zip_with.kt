// vybe-test: kotlin/kotlin_collection_chunking/test_zip_with
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3)
            val chars = listOf("a", "b", "c")
            val out = nums.zip(chars) { n, c -> "$c${'$'}n" }
            __check((out.joinToString(",")).toString(), "a1,b2,c3")
        }
