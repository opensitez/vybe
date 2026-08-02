// vybe-test: kotlin/kotlin_iterable_to_collections/test_associate_by
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("aa", "b", "ccc")
            val out = values.associateBy { it.length }
            __check((out[1]).toString(), "b")
            __check((out[2]).toString(), "aa")
            __check((out[3]).toString(), "ccc")
        }
