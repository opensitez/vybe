// vybe-test: kotlin/kotlin_pairs_apis/test_pair_zip_with_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val merged = listOf(1, 2, 3).zip(listOf("a", "b", "c")) { i, s -> s + i }
            __check((merged.joinToString(",")).toString(), "a1,b2,c3")
        }
